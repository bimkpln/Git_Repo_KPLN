import importlib.util
from pathlib import Path
import unittest
from unittest.mock import patch

import httpx


MODULE_PATH = Path(__file__).resolve().parents[1] / "server" / "navis_mcp_server.py"
spec = importlib.util.spec_from_file_location("navis_comments_test_server", MODULE_PATH)
server = importlib.util.module_from_spec(spec)
spec.loader.exec_module(server)


class CommentTransportTests(unittest.TestCase):
    def test_tool_registered_without_status_parameter(self):
        import asyncio
        tools = asyncio.run(server.mcp.list_tools())
        tool = next(item for item in tools if item.name == "batch_add_clash_result_comments")
        self.assertEqual(set(tool.inputSchema["properties"]), {"test_name", "result_names", "comment"})

    def call_with_response(self, callback, payload, status=200):
        requests = []

        def handler(request):
            requests.append(request)
            return httpx.Response(status, json=payload)

        client = httpx.Client(base_url=server.BRIDGE_URL,
                             transport=httpx.MockTransport(handler))
        self.addCleanup(client.close)
        with patch.object(server, "_client", return_value=client) as factory:
            result = callback()
            factory.assert_called_once_with(timeout=300.0)
        return result, requests

    def test_exact_text_unicode_and_single_post(self):
        import json
        text = "  КР - устранять ТОЛЬКО между профнастилом\n(его может не быть)  "
        result, requests = self.call_with_response(
            lambda: server.batch_add_clash_result_comments(
                "Тест / #1", ["Конфликт1", "Конфликт1", " Конфликт2 "], text),
            {"updated": ["Конфликт1"], "unchanged": [" Конфликт2 "], "failed": {}},
        )
        self.assertEqual(len(requests), 1)
        self.assertEqual(requests[0].method, "POST")
        self.assertEqual(requests[0].url.raw_path.decode(),
                         f"/clash/results/{server._q('Тест / #1')}/comments")
        body = json.loads(requests[0].content)
        self.assertEqual(body, {"resultNames": ["Конфликт1", " Конфликт2 "], "comment": text})
        self.assertNotIn("status", body)
        self.assertEqual(result["unchanged"], [" Конфликт2 "])

    def test_invalid_input_does_not_connect(self):
        cases = [("", ["a"], "x"), ("t", [], "x"), ("t", [""], "x"),
                 ("t", [None], "x"), ("t", "a", "x"), ("t", ["a"], " \n"),
                 ("t", ["a"], None)]
        with patch.object(server, "_client") as client:
            for args in cases:
                with self.subTest(args=args), self.assertRaises(ValueError):
                    server.batch_add_clash_result_comments(*args)
            client.assert_not_called()

    def test_partial_errors_preserve_full_detail(self):
        result, _ = self.call_with_response(
            lambda: server.batch_add_clash_result_comments("t", ["a", "b"], "x"),
            {"updated": ["a"], "unchanged": [], "failed": {"b": "Exception\nInner exception\nStack"}},
        )
        self.assertEqual(result["updated"], ["a"])
        self.assertEqual(result["failed"], [{"result_name": "b", "error": "Exception\nInner exception\nStack"}])

    def test_http_errors_are_not_empty_successes(self):
        for status in (400, 404, 409, 500):
            with self.subTest(status=status):
                with self.assertRaisesRegex(RuntimeError, f"HTTP {status}") as raised:
                    self.call_with_response(
                        lambda: server.batch_add_clash_result_comments("t", ["a"], "x"),
                        {"error": "Full error\nInner error\nStack trace"}, status,
                    )
                self.assertIn("Inner error\nStack trace", str(raised.exception))

    def test_timeout_does_not_retry(self):
        with patch.object(server, "_client") as factory:
            client = factory.return_value.__enter__.return_value
            client.post.side_effect = httpx.ReadTimeout("still running")
            with self.assertRaisesRegex(RuntimeError, "still running"):
                server.batch_add_clash_result_comments("t", ["a"], "x")
            client.post.assert_called_once()

    def test_status_comment_backwards_compatible(self):
        import json
        requests = []

        def handler(request):
            requests.append(request)
            return httpx.Response(200, json={"ok": True})

        client = httpx.Client(base_url=server.BRIDGE_URL, transport=httpx.MockTransport(handler))
        self.addCleanup(client.close)
        with patch.object(server, "_client", return_value=client):
            server.set_clash_result_status("t", "a", "New", "exact text")
        self.assertEqual(json.loads(requests[0].content), {"status": "New", "comment": "exact text"})

    def test_actual_group_target_is_exposed(self):
        target = {"kind": "group", "name": "group", "id": "guid",
                  "result_names": ["a", "b"], "comments": ["user: note"], "outcome": "updated"}
        result, _ = self.call_with_response(
            lambda: server.batch_add_clash_result_comments("t", ["a", "b"], "note"),
            {"updated": ["a", "b"], "unchanged": [], "failed": {}, "targets": [target]},
        )
        self.assertEqual(result["targets"], [target])
        self.assertEqual(len(result["updated"]), 2)
        self.assertEqual(len(result["targets"]), 1)

    def test_read_targets_is_separate_from_child_comments(self):
        import json
        target = {"kind": "group", "name": "group", "id": "guid",
                  "result_names": ["a"], "comments": ["user: group note"]}
        requests = []

        def handler(request):
            requests.append(request)
            return httpx.Response(200, json=[target])

        client = httpx.Client(base_url=server.BRIDGE_URL, transport=httpx.MockTransport(handler))
        self.addCleanup(client.close)
        with patch.object(server, "_client", return_value=client):
            result = server.get_clash_comment_targets("t", ["a", "a"])
        self.assertEqual(result["targets"], [target])
        self.assertEqual(requests[0].url.path, "/clash/results/t/comment-targets")
        self.assertEqual(json.loads(requests[0].content), {"resultNames": ["a"]})


if __name__ == "__main__":
    unittest.main()
