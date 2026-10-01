import json
from pathlib import Path
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from create_record import create_record, provenance


class RecordsTests(unittest.TestCase):
    def test_no_invented_decision_or_override(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "input.json"
            path.write_text('{"document_id":"synthetic"}', encoding="utf-8")
            record = create_record(path, {"source": "test"})
            self.assertEqual(record["decisions"], [])
            self.assertEqual(record["overrides"], [])
            self.assertEqual(record["project_context"], {"source": "test"})
            self.assertEqual(len(record["input_sha256"]), 64)

    def test_fingerprint_normalizes_newlines_and_changes_on_policy_change(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "rules").mkdir()
            (root / "VERSION").write_text("1.0.0\n")
            path = root / "rules" / "policy.md"
            path.write_bytes(b"line1\r\nline2\r\n")
            first = provenance(root)["fingerprint_sha256"]
            path.write_bytes(b"line1\nline2\n")
            self.assertEqual(provenance(root)["fingerprint_sha256"], first)
            path.write_bytes(b"changed\n")
            self.assertNotEqual(provenance(root)["fingerprint_sha256"], first)


if __name__ == "__main__": unittest.main()
