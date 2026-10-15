"""Boundary tests: reject bad geometry before HTTP and preserve the write contract."""
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "Server"))
import revit_mcp_server as server
from pydantic import ValidationError


class DocumentationTests(unittest.TestCase):
    def test_crop_copy_rejects_same_view_and_preserves_atomic_preview(self):
        data = dict(view_unique_id="target", source_view_unique_id="source")
        item = server.PlanCropCopy(**data)
        for extra in ({"view_unique_id": ""}, {"source_view_unique_id": "target"}, {"crop_min_mm": [0, 0, 0]}):
            with self.subTest(extra=extra), self.assertRaises(ValidationError):
                server.PlanCropCopy(**(data | extra))
        with patch.object(server.client, "command", return_value={}) as send:
            server.copy_revit_plan_crop("doc", 7, [item], session_id="s")
            send.assert_called_once_with("copy_plan_crop", "s", document_id="doc", expected_revision=7, items=[data], dry_run=True)
            server.copy_revit_plan_crop("doc", 8, [item], dry_run=False, session_id="s")
            self.assertFalse(send.call_args.kwargs["dry_run"])

    def test_set_plan_range_requires_finite_offsets_and_preserves_scope(self):
        data = dict(view_unique_id="trc5", top_offset_mm=4650, cut_offset_mm=1200,
                    bottom_offset_mm=-150, depth_offset_mm=-150)
        item = server.PlanRangeSet(**data)
        for key in ("top_offset_mm", "cut_offset_mm", "bottom_offset_mm", "depth_offset_mm"):
            missing = data.copy()
            missing.pop(key)
            with self.subTest(key=key), self.assertRaises(ValidationError):
                server.PlanRangeSet(**missing)
            for value in (float("nan"), float("inf"), -float("inf")):
                with self.subTest(key=key, value=value), self.assertRaises(ValidationError):
                    server.PlanRangeSet(**(data | {key: value}))
        with self.assertRaises(ValidationError):
            server.PlanRangeSet(**(data | {"view_unique_id": ""}))
        with self.assertRaises(ValidationError):
            server.PlanRangeSet(**(data | {"top_offset": 500}))
        with patch.object(server.client, "command", return_value={}) as send:
            server.set_revit_plan_range("doc", 42, [item], session_id="session")
            send.assert_called_once_with("set_plan_range", "session", document_id="doc",
                expected_revision=42, items=[data], dry_run=True)
            server.set_revit_plan_range("doc", 44, [item], dry_run=False, session_id="session")
            self.assertFalse(send.call_args.kwargs["dry_run"])

    def test_linked_rooms_batch_preserves_scope_and_preview(self):
        base = dict(view_unique_id="plan", link_instance_unique_id="link", tag_type_unique_id="room-tag")
        for extra in ({"view_unique_id": ""}, {"orientation": "Diagonal"}, {"linked_level_unique_id": ""}, {"head_mm": [0, 0, 0]}):
            with self.subTest(extra=extra), self.assertRaises(ValidationError):
                server.LinkedRoomTagCreate(**(base | extra))
        items = [server.LinkedRoomTagCreate(**base),
                 server.LinkedRoomTagCreate(**(base | {"view_unique_id": "ceiling"}), linked_level_unique_id="linked-level", orientation="Vertical")]
        with patch.object(server.client, "command", return_value={}) as send:
            server.tag_revit_linked_rooms("doc", 7, items, session_id="session")
            self.assertEqual(send.call_args.args, ("tag_linked_rooms", "session"))
            data = send.call_args.kwargs
            self.assertEqual((data["document_id"], data["expected_revision"], data["dry_run"]), ("doc", 7, True))
            self.assertNotIn("linked_level_unique_id", data["items"][0])
            self.assertEqual(data["items"][1]["linked_level_unique_id"], "linked-level")
            self.assertEqual(data["items"][1]["orientation"], "Vertical")
            server.tag_revit_linked_rooms("doc", 8, items, dry_run=False)
            self.assertFalse(send.call_args.kwargs["dry_run"])

    def test_plan_range_copy_is_scoped_and_defaults_to_rollback(self):
        item = server.PlanRangeCopy(view_unique_id="target", source_view_unique_id="source")
        with patch.object(server.client, "command", return_value={}) as send:
            server.copy_revit_plan_range("doc", 42, [item], session_id="session")
            send.assert_called_once_with("copy_plan_range", "session", document_id="doc",
                expected_revision=42, items=[{"view_unique_id": "target", "source_view_unique_id": "source"}], dry_run=True)
        with self.assertRaises(ValidationError):
            server.PlanRangeCopy(view_unique_id="target")

    def test_topology_and_range_reads_keep_document_identity(self):
        with patch.object(server.client, "command", return_value={}) as send:
            server.get_revit_mep_connections("doc", ["pipe"], session_id="session")
            send.assert_called_once_with("get_mep_connections", "session", document_id="doc", unique_ids=["pipe"])
        with patch.object(server.client, "command", return_value={}) as send:
            server.get_revit_plan_range("doc", "view", session_id="session")
            send.assert_called_once_with("get_plan_range", "session", document_id="doc", view_unique_id="view")

    def test_tag_update_preserves_identity_position_and_dry_run(self):
        item = server.TagUpdate(view_unique_id="view", tag_unique_id="existing-tag",
                                expected_head_mm=[1, 2, 3], head_mm=[4, 5, 3], elbow_mm=[4, 2, 3])
        with patch.object(server.client, "command", return_value={}) as command:
            server.update_revit_tags("doc", 9, [item], session_id="session")
            self.assertEqual(command.call_args.args, ("update_tags", "session"))
            args = command.call_args.kwargs
            self.assertTrue(args["dry_run"])
            self.assertEqual(args["expected_revision"], 9)
            self.assertEqual(args["items"][0]["tag_unique_id"], "existing-tag")
            self.assertEqual(args["items"][0]["expected_head_mm"], [1, 2, 3])
            self.assertNotIn("leader", args["items"][0])
        with self.assertRaises(ValidationError):
            server.TagUpdate(view_unique_id="view", tag_unique_id="tag", expected_head_mm=[0, 0, 0],
                             head_mm=[1, 1, 0], leader=False, elbow_mm=[0, 1, 0])
        with self.assertRaises(ValidationError):
            server.TagUpdate(view_unique_id="view", tag_unique_id="tag", head_mm=[1, 1, 0])

    def test_placement_requires_one_source_and_sheet_coordinates(self):
        base = dict(sheet_unique_id="sheet", view_unique_id="view")
        for extra in ({}, {"point_mm": [1, 2, 0], "reference_viewport_unique_id": "sample"},
                      {"point_mm": [1, 2, 3]}, {"point_mm": [1, float("nan"), 0]}):
            with self.subTest(extra=extra), self.assertRaises(ValidationError):
                server.ViewPlacement(**base, **extra)
        self.assertEqual(server.ViewPlacement(**base, point_mm=[105, 210, 0]).point_mm, (105, 210, 0))

    def test_section_rejects_invalid_basis_and_reversed_box(self):
        base = dict(view_type_unique_id="type", name="Section", origin_mm=[0, 0, 0], up=[0, 0, 1],
                    direction=[0, 1, 0], min_mm=[-1000, 0, 0], max_mm=[1000, 3000, 500])
        for extra in ({"up": [0, 0, 2]}, {"direction": [0, 0, 1]}, {"max_mm": [1000, 3000, -1]}):
            with self.subTest(extra=extra), self.assertRaises(ValidationError):
                server.SectionCreate(**(base | extra))
        server.SectionCreate(**base)

    def test_dimension_and_crop_contracts(self):
        base = dict(view_unique_id="view", dimension_type_unique_id="type", start_mm=[0, 0, 0], end_mm=[100, 0, 0])
        for refs in (["a"], ["a", "a"]):
            with self.assertRaises(ValidationError):
                server.DimensionCreate(**base, references=refs)
        with self.assertRaises(ValidationError):
            server.PlanCreate(view_type_unique_id="type", level_unique_id="level", name="Plan", crop_min_mm=[0, 0, 0])

    def test_dry_run_is_default_and_points_are_json_arrays(self):
        item = server.TagCreate(view_unique_id="view", tag_type_unique_id="type", element_unique_id="element",
                                head_mm=[10, 20, 30], link_instance_unique_id="link")
        with patch.object(server.client, "command", return_value={"ok": True}) as send:
            server.create_revit_tags("document", 12, [item], session_id="session")
            args, payload = send.call_args
            self.assertEqual(args, ("create_tags", "session"))
            self.assertEqual(payload["document_id"], "document")
            self.assertEqual(payload["expected_revision"], 12)
            self.assertTrue(payload["dry_run"])
            self.assertEqual(payload["items"][0]["head_mm"], [10, 20, 30])
            self.assertEqual(payload["items"][0]["link_instance_unique_id"], "link")

    def test_read_tools_preserve_document_scope_and_pagination(self):
        with patch.object(server.client, "command", return_value={"ok": True}) as send:
            server.get_revit_view_annotations("document", "view", offset=200, limit=50, session_id="session")
            send.assert_called_once_with("get_view_annotations", "session", document_id="document", view_unique_id="view", offset=200, limit=50)


if __name__ == "__main__":
    unittest.main()
