"""Launch the actual MCP process and negotiate with the installed SDK client."""
import unittest
import sys
from pathlib import Path
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


class StdioTests(unittest.IsolatedAsyncioTestCase):
    async def test_initialize_tools_and_sessions(self):
        server = Path(__file__).resolve().parents[1] / "server/revit_mcp_server.py"
        async with stdio_client(StdioServerParameters(command=sys.executable, args=["-B", str(server)])) as (reader, writer):
            async with ClientSession(reader, writer) as session:
                info = await session.initialize()
                self.assertEqual(info.serverInfo.name, "KPLN_RevitMcpBridge")
                tools = await session.list_tools()
                self.assertEqual(len(tools.tools), 50)
                by_name = {tool.name: tool for tool in tools.tools}
                for name in ("update_revit_viewport_labels", "create_revit_sheets", "place_revit_views", "move_revit_sheet_items", "create_revit_tags", "tag_revit_linked_rooms", "create_revit_dimensions", "create_revit_plan_views", "create_revit_sections", "set_revit_plan_range"):
                    self.assertFalse(by_name[name].annotations.readOnlyHint)
                    self.assertFalse(by_name[name].annotations.idempotentHint)
                    self.assertTrue(by_name[name].inputSchema["properties"]["dry_run"]["default"])
                    self.assertIn("expected_revision", by_name[name].inputSchema["required"])
                for name in ("get_revit_group_members", "get_revit_sheet_contents", "get_revit_view_elements", "get_revit_view_visibility", "get_revit_schedule_data", "check_revit_intersections", "check_revit_sheet_layout"):
                    self.assertTrue(by_name[name].annotations.readOnlyHint)
                self.assertNotIn("open_revit_interference_check", by_name)
                self.assertFalse(by_name["export_revit_sheets_pdf"].annotations.readOnlyHint)
                self.assertFalse(by_name["export_revit_sheets_pdf"].annotations.destructiveHint)
                self.assertFalse(by_name["export_revit_sheets_pdf"].annotations.idempotentHint)
                self.assertFalse(by_name["print_revit_sheets_pdf"].annotations.readOnlyHint)
                self.assertFalse(by_name["print_revit_sheets_pdf"].annotations.idempotentHint)
                self.assertTrue(by_name["print_revit_sheets_pdf"].inputSchema["properties"]["dry_run"]["default"])
                self.assertNotIn("left_category_ids", by_name["check_revit_intersections"].inputSchema["properties"])
                self.assertIn("time_budget_seconds", by_name["check_revit_intersections"].inputSchema["properties"])
                family_tool = next(tool for tool in tools.tools if tool.name == "create_revit_family_types")
                self.assertIn("base_type_name", family_tool.inputSchema["properties"])
                family_set = next(tool for tool in tools.tools if tool.name == "set_revit_family_type_parameters")
                self.assertIn("types", family_set.inputSchema["properties"])
                result = await session.call_tool("list_revit_sessions", {})
                self.assertFalse(result.isError)
                self.assertIn("sessions", result.content[0].text)
                # Schema validation rejects accidental unbounded requests before HTTP.
                bad = await session.call_tool("find_revit_elements", {"document_id": "fake", "limit": 201})
                self.assertTrue(bad.isError)
                for name, args in (
                    ("update_revit_viewport_labels", {"document_id": "fake", "expected_revision": 0, "items": []}),
                    ("move_revit_sheet_items", {"document_id": "fake", "expected_revision": 0, "items": []}),
                    ("check_revit_sheet_layout", {"document_id": "fake", "sheet_unique_id": "s", "minimum_gap_mm": -1}),
                    ("tag_revit_linked_rooms", {"document_id": "fake", "expected_revision": 0, "items": []}),
                    ("tag_revit_linked_rooms", {"document_id": "fake", "expected_revision": -1, "items": [{"view_unique_id": "v", "link_instance_unique_id": "l", "tag_type_unique_id": "t"}]}),
                    ("tag_revit_linked_rooms", {"document_id": "fake", "expected_revision": 0, "items": [{"view_unique_id": "v", "tag_type_unique_id": "t"}]}),
                    ("get_revit_schedule_data", {"document_id": "fake", "schedule_unique_id": "fake", "column_limit": 51}),
                    ("get_revit_schedule_data", {"document_id": "fake", "schedule_unique_id": "fake", "section": "invalid"}),
                    ("check_revit_intersections", {"document_id": "fake", "time_budget_seconds": 61}),
                    ("check_revit_intersections", {"document_id": "fake", "max_pairs": 5001}),
                    ("export_revit_sheets_pdf", {"document_id": "fake"}),
                    ("export_revit_sheets_pdf", {"document_id": "fake", "expected_revision": -1}),
                    ("print_revit_sheets_pdf", {"document_id": "fake"}),
                    ("print_revit_sheets_pdf", {"document_id": "fake", "expected_revision": -1}),
                    ("print_revit_sheets_pdf", {"document_id": "fake", "expected_revision": 0, "timeout_seconds": 61}),
                    ("print_revit_sheets_pdf", {"document_id": "fake", "expected_revision": 0, "settings": {"isDWGExport": True}}),
                ):
                    self.assertTrue((await session.call_tool(name, args)).isError)


if __name__ == "__main__": unittest.main()

