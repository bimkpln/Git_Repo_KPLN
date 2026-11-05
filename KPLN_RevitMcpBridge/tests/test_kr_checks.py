import sys
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "Server"))
import revit_mcp_server as server


class KRChecksTests(unittest.TestCase):
    def test_catalog_uses_selected_document_and_session(self):
        with patch.object(server.client, "command", return_value={}) as send:
            server.get_revit_kr_check_catalog("doc", "session")
            send.assert_called_once_with("get_kr_check_catalog", "session", document_id="doc")

    def test_both_checks_keep_revision_and_module_and_do_not_retry_pending(self):
        for check_id in ("CheckIfcRebarMarks", "CheckRebarConcreteClass"):
            pending = {"ok": False, "operation_id": "op", "error": {"code": "outcome_pending"}}
            with patch.object(server.client, "command", return_value=pending) as send:
                self.assertIs(server.run_revit_kr_check("doc", 7, check_id, "module", "session"), pending)
                send.assert_called_once_with("run_kr_check", "session", document_id="doc",
                    expected_revision=7, check_id=check_id, library_module_id="module")

    def test_all_sections_are_read_without_run(self):
        for section in ("issues", "skipped_elements", "run_warnings"):
            with patch.object(server.client, "command", return_value={}) as send:
                server.get_revit_kr_check_result("doc", "report", section, 200, 50, "session")
                send.assert_called_once_with("get_kr_check_result", "session", document_id="doc",
                    result_id="report", section=section, offset=200, limit=50)
