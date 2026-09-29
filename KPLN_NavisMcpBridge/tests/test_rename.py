import asyncio
import importlib.util
import json
from pathlib import Path
import unittest
from unittest.mock import patch

import httpx

spec = importlib.util.spec_from_file_location(
    "navis_rename_server", Path(__file__).resolve().parents[1] / "server" / "navis_mcp_server.py")
server = importlib.util.module_from_spec(spec)
spec.loader.exec_module(server)


class RenameTransportTests(unittest.TestCase):
    def test_registered_preview_default(self):
        tool = next(t for t in asyncio.run(server.mcp.list_tools()) if t.name == "batch_rename_clash_tests")
        self.assertIs(tool.inputSchema["properties"]["dry_run"]["default"], True)

    def test_exact_names_single_post_preview_and_apply(self):
        updates = [{"old_name": " АР / #1 ", "new_name": "АРХИВ_ АР / #1 "}]
        for dry in (True, False):
            requests = []
            def handler(req):
                requests.append(req)
                return httpx.Response(200, json={"dry_run": dry, "updates": [{"id": "test-id", "outcome": "planned" if dry else "updated"}]})
            client = httpx.Client(base_url=server.BRIDGE_URL, transport=httpx.MockTransport(handler))
            self.addCleanup(client.close)
            with patch.object(server, "_client", return_value=client):
                result = server.batch_rename_clash_tests(updates, dry)
            self.assertEqual(len(requests), 1)
            self.assertEqual(requests[0].method, "POST")
            self.assertEqual(requests[0].url.path, "/clash/tests/rename")
            self.assertEqual(json.loads(requests[0].content), {"updates": [{"oldName": " АР / #1 ", "newName": "АРХИВ_ АР / #1 "}], "dryRun": dry})
            self.assertEqual(result["updates"][0]["id"], "test-id")

    def test_invalid_input_never_connects(self):
        cases = [None, [], "a", [None], [{}], [{"old_name": "a", "new_name": " "}],
                 [{"old_name": "a", "new_name": 1}], [{"old_name": "a", "new_name": "x", "status": "Approved"}],
                 [{"old_name": "a", "new_name": "x"}, {"old_name": "a", "new_name": "y"}],
                 [{"old_name": "a", "new_name": "x"}, {"old_name": "b", "new_name": "x"}]]
        with patch.object(server, "_client") as factory:
            for case in cases:
                with self.subTest(case=case), self.assertRaises(ValueError):
                    server.batch_rename_clash_tests(case)
            with self.assertRaises(ValueError):
                server.batch_rename_clash_tests([{"old_name": "a", "new_name": "b"}], "false")
            factory.assert_not_called()

    def test_http_error_preserved(self):
        for status in (400, 404, 409, 500):
            client = httpx.Client(base_url=server.BRIDGE_URL, transport=httpx.MockTransport(
                lambda req: httpx.Response(status, json={"error": "detailed rename failure"})))
            self.addCleanup(client.close)
            with patch.object(server, "_client", return_value=client), self.assertRaisesRegex(RuntimeError, f"HTTP {status}"):
                server.batch_rename_clash_tests([{"old_name": "a", "new_name": "b"}], False)

    def test_timeout_never_retries(self):
        with patch.object(server, "_client") as factory:
            client = factory.return_value.__enter__.return_value
            client.post.side_effect = httpx.ReadTimeout("unknown outcome")
            with self.assertRaises(RuntimeError):
                server.batch_rename_clash_tests([{"old_name": "a", "new_name": "b"}], False)
            client.post.assert_called_once()


if __name__ == "__main__":
    unittest.main()
