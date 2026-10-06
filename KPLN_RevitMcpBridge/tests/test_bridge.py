import asyncio
import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

import httpx
from pydantic import ValidationError

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "server"))
from bridge_client import BridgeClient, BridgeError
import revit_mcp_server as server


class BridgeTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.directory = Path(self.temp.name)
        self.session_id = "a" * 32
        self.data = dict(session_id=self.session_id, url="http://127.0.0.1:18765", token="T" * 44, protocol=1, pid=123, revit_version="2024")
        self.save()
        self.requests = []
        self.client = BridgeClient(self.directory, httpx.MockTransport(self.handler))

    def tearDown(self):
        self.temp.cleanup()

    def save(self):
        (self.directory / (self.data["session_id"] + ".json")).write_text(json.dumps(self.data), encoding="utf-8")

    def handler(self, request):
        self.requests.append(request)
        self.assertEqual(request.headers["Authorization"], "Bearer " + "T" * 44)
        if request.url.path == "/health":
            return httpx.Response(200, json=dict(ok=True, session_id=self.session_id, protocol=1))
        return httpx.Response(200, json=dict(ok=True, result={"dry_run": True}))

    def test_sessions_never_expose_token(self):
        result = self.client.sessions()
        self.assertEqual(len(result["sessions"]), 1)
        self.assertNotIn("token", json.dumps(result))
        self.assertNotIn("T" * 44, json.dumps(result))

    def test_command_has_identity_and_no_retry(self):
        self.client.command("get_context", self.session_id)
        self.assertEqual(len(self.requests), 1)
        payload = json.loads(self.requests[0].content)
        self.assertEqual(payload["session_id"], self.session_id)
        self.assertEqual(len(payload["operation_id"]), 32)

    def test_reject_remote_discovery(self):
        for url in ["http://example.com:18765", "https://127.0.0.1:18765", "http://127.0.0.1:80", "http://user:pass@127.0.0.1:18765", "http://127.0.0.1:18765/path"]:
            self.data["url"] = url
            self.save()
            self.assertEqual(self.client.sessions()["sessions"], [])
        self.assertEqual(len(self.requests), 0)

    def test_path_traversal_rejected(self):
        with self.assertRaises(BridgeError): self.client.health("../../secret")

    def test_wrong_session_health_rejected(self):
        self.session_id = "b" * 32
        self.assertEqual(self.client.sessions()["sessions"], [])

    def test_no_session_actionable(self):
        for path in self.directory.glob("*.json"): path.unlink()
        with self.assertRaisesRegex(BridgeError, "доступно сессий 0"): self.client.command("get_context")

    def test_multiple_sessions_require_choice(self):
        with patch.object(self.client, "sessions", return_value={"sessions": [{}, {}]}):
            with self.assertRaisesRegex(BridgeError, "session_id"): self.client.command("get_context")

    def test_timeout_exposes_operation_without_retry(self):
        requests = []
        def timeout(request):
            requests.append(request)
            raise httpx.ReadTimeout("timeout", request=request)
        self.client.transport = httpx.MockTransport(timeout)
        with self.assertRaisesRegex(BridgeError, "operation_id="):
            self.client.command("set_parameters", self.session_id, dry_run=False)
        self.assertEqual(len(requests), 1)

    def test_bridge_error_propagates(self):
        self.client.transport = httpx.MockTransport(lambda r: httpx.Response(409, json={"ok": False, "error": {"code": "stale_document", "message": "changed"}}))
        with self.assertRaisesRegex(BridgeError, "stale_document"):
            self.client.command("set_parameters", self.session_id)

    def test_redirect_not_followed(self):
        self.client.transport = httpx.MockTransport(lambda r: httpx.Response(302, headers={"location": "http://example.com"}))
        with self.assertRaisesRegex(BridgeError, "перенаправление"): self.client.health(self.session_id)

    def test_malformed_json_handled(self):
        self.client.transport = httpx.MockTransport(lambda r: httpx.Response(200, text="<html>bad</html>"))
        with self.assertRaisesRegex(BridgeError, "JSON"): self.client.health(self.session_id)

    def test_operation_path_normalized(self):
        self.client.operation("BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB", self.session_id)
        self.assertEqual(self.requests[-1].url.path, "/operations/" + "b" * 32)

    def test_set_defaults_to_preview(self):
        update = server.ParameterUpdate(element_unique_id="x", parameter_id="-1", expected_value="before", value="after")
        with patch.object(server.client, "command", return_value={}) as command:
            server.set_revit_parameters("doc", 1, [update])
        self.assertTrue(command.call_args.kwargs["dry_run"])

    def test_expected_value_is_required(self):
        with self.assertRaises(ValidationError): server.ParameterUpdate(element_unique_id="x", parameter_id="-1", value="after")
        with self.assertRaises(ValidationError): server.FamilyValueSet(parameter_id="-1", value="after")

    def test_family_set_defaults_to_preview(self):
        value = server.FamilyValueSet(parameter_id="-1", expected_value="before", value="after")
        family_type = server.FamilyTypeSet(name="Type A", values=[value])
        with patch.object(server.client, "command", return_value={}) as command:
            server.set_revit_family_type_parameters("doc", 1, [family_type])
        self.assertEqual(command.call_args.args[0], "set_family_type_parameters")
        self.assertTrue(command.call_args.kwargs["dry_run"])

    def test_boolean_not_numeric_value(self):
        with self.assertRaises(ValidationError): server.ParameterUpdate(element_unique_id="x", parameter_id="-1", expected_value=1, value=True)

    def test_tool_schemas_and_annotations(self):
        tools = asyncio.run(server.mcp.list_tools())
        self.assertEqual(len(tools), 27)
        for tool in tools:
            self.assertIsNotNone(tool.annotations)
            self.assertFalse(tool.annotations.openWorldHint)
        write = next(t for t in tools if t.name == "set_revit_parameters")
        self.assertFalse(write.annotations.readOnlyHint)
        self.assertTrue(write.inputSchema["properties"]["dry_run"]["default"])
        family_write = next(t for t in tools if t.name == "create_revit_family_types")
        self.assertFalse(family_write.annotations.readOnlyHint)
        self.assertTrue(family_write.inputSchema["properties"]["dry_run"]["default"])
        family_set = next(t for t in tools if t.name == "set_revit_family_type_parameters")
        self.assertFalse(family_set.annotations.readOnlyHint)
        self.assertTrue(family_set.inputSchema["properties"]["dry_run"]["default"])

    def test_inspection_tools_forward_identity_and_scope(self):
        cases = (
            (server.get_revit_group_members, {"group_unique_id": "group", "offset": 100}, "get_group_members"),
            (server.get_revit_sheet_contents, {"sheet_unique_id": "sheet"}, "get_sheet_contents"),
            (server.get_revit_view_elements, {"view_unique_id": "view", "category_id": "-2000011"}, "get_view_elements"),
            (server.get_revit_view_visibility, {"view_unique_id": "view", "unique_ids": ["wall"]}, "get_view_visibility"),
            (server.get_revit_schedule_data, {"schedule_unique_id": "schedule", "column_offset": 50}, "get_schedule_data"),
            (server.check_revit_intersections, {"left_category_ids": ["-2000011"], "right_category_ids": ["-2000032"]}, "check_intersections"),
            (server.open_revit_interference_check, {}, "open_interference_check"),
            (server.export_revit_sheets_pdf, {"expected_revision": 2}, "export_sheets_pdf"),
        )
        for function, args, command_name in cases:
            with self.subTest(command=command_name), patch.object(server.client, "command", return_value={}) as command:
                function(document_id="doc", session_id="session", **args)
                self.assertEqual(command.call_args.args, (command_name, "session"))
                self.assertEqual(command.call_args.kwargs["document_id"], "doc")
                for key, value in args.items():
                    self.assertEqual(command.call_args.kwargs[key], value)

    def test_pdf_export_defaults_to_preview_and_preserves_revision(self):
        with patch.object(server.client, "command", return_value={}) as command:
            server.export_revit_sheets_pdf("doc", 42, session_id="session")
        self.assertTrue(command.call_args.kwargs["dry_run"])
        self.assertEqual(command.call_args.kwargs["expected_revision"], 42)

    def test_publication_print_defaults_match_requested_setup(self):
        with patch.object(server.client, "command", return_value={}) as command:
            server.print_revit_sheets_pdf("doc", 42, session_id="session")
        self.assertEqual(command.call_args.args, ("print_sheets_pdf", "session"))
        self.assertTrue(command.call_args.kwargs["dry_run"])
        self.assertEqual(command.call_args.kwargs["expected_revision"], 42)
        settings = command.call_args.kwargs["settings"]
        self.assertEqual(settings["printerName"], "PDFCreator")
        self.assertEqual(settings["outputPDFFolder"], "C:\\PDF_Print")
        self.assertEqual(settings["pdfNameConstructor"], "<Номер листа>_<Имя листа>.pdf")
        self.assertEqual(settings["colorsType"], "Color")
        self.assertEqual(settings["hiddenLineProcessing"], "VectorProcessing")
        self.assertTrue(settings["isRefreshSchedules"])
        for name in ("isDWGExport", "isPrintToPaper", "isMergePdfs", "isExcludeBorders", "isUseOrientation"):
            self.assertFalse(settings[name])

    def test_publication_unsupported_modes_are_rejected(self):
        for value in ({"printerName": "Kyocera"}, {"isPrintToPaper": True}, {"isDWGExport": True},
                      {"isMergePdfs": True}, {"isExcludeBorders": True}, {"unknown": True}):
            with self.subTest(value=value), self.assertRaises(ValidationError):
                server.BatchPrintSettings(**value)


if __name__ == "__main__": unittest.main()
