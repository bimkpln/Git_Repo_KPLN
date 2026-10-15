"""MCP boundary and transport contracts for existing sheet items."""
import sys
import unittest
from pathlib import Path
from unittest.mock import patch
from pydantic import ValidationError

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "Server"))
import revit_mcp_server as server


class SheetLayoutTests(unittest.TestCase):
    def test_label_offset_guards_and_transport(self):
        base = dict(sheet_unique_id="sheet", viewport_unique_id="viewport",
                    expected_label_offset_mm=[3, 3, 0], label_offset_mm=[200, 400, 0])
        for changes in ({"sheet_unique_id": ""}, {"viewport_unique_id": ""},
                        {"label_offset_mm": [1, 2, 3]}, {"expected_label_offset_mm": [1, 2, 3]},
                        {"label_offset_mm": [float("nan"), 0, 0]},
                        {"label_line_length_mm": 5}, {"expected_label_line_length_mm": 5},
                        {"label_line_length_mm": float("inf"), "expected_label_line_length_mm": 5},
                        {"ignore_overlap": True}):
            with self.subTest(changes=changes), self.assertRaises(ValidationError):
                server.ViewportLabelUpdate(**(base | changes))
        with self.assertRaises(ValidationError):
            server.ViewportLabelUpdate(**{k: v for k, v in base.items() if k != "expected_label_offset_mm"})
        item = server.ViewportLabelUpdate(**base, expected_label_line_length_mm=170, label_line_length_mm=-120)
        with patch.object(server.client, "command", return_value={}) as send:
            server.update_revit_viewport_labels("doc", 7, [item], session_id="session")
            self.assertEqual(send.call_args.args, ("update_viewport_labels", "session"))
            self.assertTrue(send.call_args.kwargs["dry_run"])
            self.assertEqual(send.call_args.kwargs["items"][0]["label_line_length_mm"], -120)
            server.update_revit_viewport_labels("doc", 9, [item], dry_run=False)
            self.assertFalse(send.call_args.kwargs["dry_run"])
            with self.assertRaises(ValueError):
                server.update_revit_viewport_labels("doc", 9, [item, item])

    def test_move_requires_observed_position_and_planar_finite_points(self):
        base = dict(sheet_unique_id="sheet", instance_unique_id="schedule",
                    expected_point_mm=[-200, 560, 0], point_mm=[-139, 560, 0])
        for patch_value in ({"sheet_unique_id": ""}, {"instance_unique_id": ""},
                            {"expected_point_mm": [1, 2, 17]}, {"point_mm": [1, 2, 1]},
                            {"point_mm": [float("nan"), 0, 0]}, {"point_mm": [0, float("inf"), 0]},
                            {"ignore_overlap": True}):
            with self.subTest(patch_value=patch_value), self.assertRaises(ValidationError):
                server.SheetItemMove(**(base | patch_value))
        without_expected = base.copy()
        without_expected.pop("expected_point_mm")
        with self.assertRaises(ValidationError):
            server.SheetItemMove(**without_expected)
        item = server.SheetItemMove(**base)
        with patch.object(server.client, "command", return_value={}) as send:
            server.move_revit_sheet_items("doc", 42, [item], session_id="session")
            send.assert_called_once_with("move_sheet_items", "session", document_id="doc",
                                         expected_revision=42, items=[base], dry_run=True)
            server.move_revit_sheet_items("doc", 44, [item], dry_run=False, session_id="session")
            self.assertFalse(send.call_args.kwargs["dry_run"])
            with self.assertRaises(ValueError):
                server.move_revit_sheet_items("doc", 45, [item, item])

    def test_frame_and_reserved_fields_are_explicit_and_validated(self):
        with patch.object(server.client, "command", return_value={}) as send:
            for args in ({"content_min_mm": [-800, 5, 0]},
                         {"content_min_mm": [-800, 5, 0], "content_max_mm": [-810, 590, 0]},
                         {"content_min_mm": [-800, 5, 0], "content_max_mm": [-5, 590, 7]}):
                with self.subTest(args=args), self.assertRaises(ValueError):
                    server.check_revit_sheet_layout("doc", "sheet", **args)
            send.assert_not_called()
            reserved = server.SheetReservedRegion(name="stamp", min_mm=[-190, 5, 0], max_mm=[-5, 65, 0])
            server.check_revit_sheet_layout("doc", "sheet", minimum_gap_mm=2,
                                           content_min_mm=[-821, 5, 0], content_max_mm=[-5, 589, 0],
                                           reserved_regions=[reserved], session_id="session")
            self.assertEqual(send.call_args.args, ("check_sheet_layout", "session"))
            self.assertEqual(send.call_args.kwargs["reserved_regions"][0]["name"], "stamp")
            self.assertNotIn("expected_revision", send.call_args.kwargs)


if __name__ == "__main__":
    unittest.main()
