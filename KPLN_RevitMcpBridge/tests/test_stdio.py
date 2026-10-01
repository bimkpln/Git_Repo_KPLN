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
                self.assertEqual(len(tools.tools), 18)
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


if __name__ == "__main__": unittest.main()
