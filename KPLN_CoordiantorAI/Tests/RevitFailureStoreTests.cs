using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using KPLN_CoordiantorAI.ExternalAIModel.Mcp;

internal static class RevitFailureStoreTests
{
    private static int _checks;
    private static readonly string Root = Path.Combine(Path.GetTempPath(), "CoordinatorFailureTests_" + Guid.NewGuid().ToString("N"));
    private static void Check(bool condition, string label)
    {
        if (!condition) throw new Exception(label);
        _checks++;
    }
    private static RevitFailureStore NewStore(string name)
    {
        return new RevitFailureStore(2023, (m, e) => { }, Path.Combine(Root, name));
    }
    private static JObject Record(string document = "a", int second = 0, string severity = "Error", string text = "Test failure")
    {
        return new JObject
        {
            { "document_id", document }, { "document_name", "Same model name" },
            { "observed_at_utc", new DateTimeOffset(2026, 10, 15, 8, 0, 0, TimeSpan.Zero).AddSeconds(second).ToString("o") },
            { "severity", severity }, { "description", text }, { "element_ids", new JArray(123) }
        };
    }
    private static JObject Read(RevitFailureStore store, JObject args = null)
    {
        var response = store.Execute(args ?? new JObject(), CancellationToken.None);
        if (!response.Success) throw new Exception(response.Error.Code + ": " + response.Error.Message);
        return (JObject)response.Result;
    }
    private static bool HasSnapshot(string path, int count)
    {
        try
        {
            // Observe the background writer without taking a lock that denies its next write.
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var reader = new StreamReader(stream))
            {
                string[] lines = reader.ReadToEnd().Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                return lines.Length == count && lines.All(l => JObject.Parse(l)["description"] != null);
            }
        }
        catch { return false; }
    }
    private static void Main()
    {
        Directory.CreateDirectory(Root);
        try
        {
            FilteringAndPagination();
            BoundsAndLifecycle();
            CleanupAndWriteFailure();
            HttpWhileRevitIsBusy().GetAwaiter().GetResult();
            Console.WriteLine("PASS: " + _checks + " checks (store, lifecycle, schema and real HTTP transport).");
        }
        finally
        {
            // Root is a fresh test-owned GUID directory, never the user's failure directory.
            if (Directory.Exists(Root)) Directory.Delete(Root, true);
        }
    }

    private static void FilteringAndPagination()
    {
        using (var store = NewStore("filter"))
        {
            Check(!File.Exists(store.FilePath), "No file before a failure");
            Check(store.Execute(new JObject(), CancellationToken.None).Error.Code == "active_document_unknown", "Unknown active document");
            store.SetActiveDocument("a", "Same model name");
            Check((int)Read(store)["total_count"] == 0, "Empty current document");
            for (int i = 0; i < 25; i++) store.Add(Record("a", i));
            store.Add(Record("b", 0, "Warning"));
            JObject first = Read(store);
            Check((int)first["total_count"] == 25 && ((JArray)first["failures"]).Count == 10, "Default ten/current scope");
            Check((long)first["failures"][0]["sequence"] == 25, "Newest first");
            store.Add(Record("a", 30));
            JObject page2 = Read(store, new JObject { { "offset", 10 }, { "snapshot_sequence", first["snapshot_sequence"] } });
            Check((long)page2["failures"][0]["sequence"] == 15, "Stable upper bound when new failures arrive");
            JObject last = Read(store, new JObject { { "offset", 20 }, { "snapshot_sequence", first["snapshot_sequence"] } });
            Check((int)last["returned_count"] == 5 && !(bool)last["has_more"] && last["next_offset"].Type == JTokenType.Null, "Final page");
            Check((int)Read(store, new JObject { { "all_documents", true } })["total_count"] == 27, "All documents");
            Check((int)Read(store, new JObject { { "document_id", "b" } })["total_count"] == 1, "Same model names remain isolated");
            Check((int)Read(store, new JObject { { "dateTime", "2026-10-15 11:00:05" }, { "endDateTime", "2026-10-15T11:00:10" } })["total_count"] == 6, "Moscow inclusive bounds");
            Check((int)Read(store, new JObject { { "dateTime", "2026-10-15T08:00:05Z" }, { "endDateTime", "2026-10-15T11:00:10+03:00" } })["total_count"] == 6, "Explicit UTC and offsets");
            Check((int)Read(store, new JObject { { "all_documents", true }, { "severity", "Warning" } })["total_count"] == 1, "Severity");
            JObject[] invalid = {
                new JObject { { "limit", 0 } }, new JObject { { "limit", 201 } },
                new JObject { { "limit", "10" } }, new JObject { { "offset", -1 } },
                new JObject { { "dateTime", "yesterday" } },
                new JObject { { "dateTime", "2026-10-16 00:00:00" }, { "endDateTime", "2026-10-15 00:00:00" } },
                new JObject { { "severity", "unknown" } }, new JObject { { "all_documents", "true" } },
                new JObject { { "document_id", "a" }, { "all_documents", true } }, new JObject { { "unknown", 1 } }
            };
            foreach (JObject args in invalid)
                Check(store.Execute(args, CancellationToken.None).Error.Code == "invalid_arguments", "Invalid arguments");
            Check(store.Execute(new JObject { { "document_id", "missing" } }, CancellationToken.None).Error.Code == "document_not_found", "Unknown document");
            Check((int)Read(store, new JObject { { "offset", 1000 } })["returned_count"] == 0, "Offset beyond end");
            using (var cancel = new CancellationTokenSource())
            {
                cancel.Cancel();
                bool canceled = false;
                try { store.Execute(new JObject(), cancel.Token); } catch (OperationCanceledException) { canceled = true; }
                Check(canceled, "Cancellation");
            }
        }
    }

    private static void BoundsAndLifecycle()
    {
        var store = NewStore("bounds");
        store.SetActiveDocument("a", "Model");
        for (int i = 0; i < 205; i++) store.Add(Record("a", i));
        JObject result = Read(store, new JObject { { "limit", 200 } });
        Check((int)result["total_count"] == 200 && (long)result["dropped_count"] == 5, "Per-document retention");
        Check(SpinWait.SpinUntil(() => HasSnapshot(store.FilePath, 200), 5000), "Bounded JSONL snapshot");
        string path = store.FilePath;
        store.Dispose();
        store.Add(Record());
        Check(!File.Exists(path), "Dispose deletes the session file");
        Check(store.Execute(new JObject(), CancellationToken.None).Error.Code == "failure_capture_stopped", "Stopped store");

        using (var heavy = NewStore("heavy"))
        {
            heavy.SetActiveDocument("a", "Model");
            for (int i = 0; i < 30; i++) heavy.Add(Record("a", i, "Error", new string('\u0416', 8192)));
            JObject page = Read(heavy, new JObject { { "limit", 200 } });
            Check((int)page["returned_count"] < 30 && (bool)page["has_more"], "Byte-limited page");
            Check(Encoding.UTF8.GetByteCount(JsonConvert.SerializeObject(page, Formatting.None)) <= RevitFailureStore.PageBytes, "Actual UTF-8 page size");
            int count = 0;
            int offset = 0;
            do
            {
                page = Read(heavy, new JObject { { "limit", 200 }, { "offset", offset } });
                count += (int)page["returned_count"];
                offset += (int)page["returned_count"];
            } while ((bool)page["has_more"]);
            Check(count == 30, "All large records reachable");
            for (int i = 0; i < 600; i++) heavy.Add(Record("doc" + (i % 4), i, "Error", new string('x', 8192)));
            Check((bool)Read(heavy, new JObject { { "all_documents", true } })["buffer_incomplete"], "Global byte budget reports eviction");
        }

        for (int i = 0; i < 20; i++)
        {
            var race = NewStore("race" + i);
            race.Add(Record());
            race.Dispose();
            Check(!File.Exists(race.FilePath), "Writer cannot recreate deleted file");
        }
    }

    private static void CleanupAndWriteFailure()
    {
        string dir = Path.Combine(Root, "cleanup");
        Directory.CreateDirectory(dir);
        using (var process = Process.GetCurrentProcess())
        {
            string active = Path.Combine(dir, "failures_2024_" + process.Id + "_" +
                process.StartTime.ToUniversalTime().Ticks + "_" + Guid.NewGuid().ToString("N") + ".jsonl");
            string oldPid = Path.Combine(dir, "failures_2023_" + process.Id + "_1_" + Guid.NewGuid().ToString("N") + ".jsonl");
            string dead = Path.Combine(dir, "failures_2023_2147483647_1_" + Guid.NewGuid().ToString("N") + ".jsonl");
            string unrelated = Path.Combine(dir, "failures_notes.jsonl");
            foreach (string p in new[] { active, oldPid, dead, unrelated }) File.WriteAllText(p, "test");
            RevitFailureStore.CleanupAbandonedFiles(dir, null);
            Check(File.Exists(active), "Live process file preserved");
            Check(!File.Exists(oldPid) && !File.Exists(dead), "Dead/reused process files removed");
            Check(File.Exists(unrelated), "Unrecognized files preserved");
        }
        string blocked = Path.Combine(Root, "not-a-directory");
        File.WriteAllText(blocked, "block");
        using (var store = new RevitFailureStore(2023, null, blocked))
        {
            store.SetActiveDocument("a", "Model");
            store.Add(Record());
            Check(SpinWait.SpinUntil(() => Read(store)["file_write_error"].Type != JTokenType.Null, 5000), "Write error surfaced");
            Check((int)Read(store)["total_count"] == 1, "Memory survives disk errors");
        }
    }

    private static async Task HttpWhileRevitIsBusy()
    {
        ServicePointManager.DefaultConnectionLimit = 8;
        int port;
        var tcp = new TcpListener(IPAddress.Loopback, 0);
        tcp.Start();
        port = ((IPEndPoint)tcp.LocalEndpoint).Port;
        tcp.Stop();
        string endpoint = "http://127.0.0.1:" + port + "/mcp/";
        using (var store = NewStore("http"))
        using (var client = new HttpClient(new HttpClientHandler { UseProxy = false }) { Timeout = TimeSpan.FromSeconds(5) })
        {
            store.SetActiveDocument("a", "Model");
            store.Add(Record());
            var handler = new RevitMcpExternalEventHandler();
            var transport = new RevitMcpTransport(endpoint, handler, "test", 2023, store);
            transport.Start();
            Task<HttpResponseMessage> pending = null;
            try
            {
                pending = client.PostAsync(endpoint, Body("get_active_view_in_revit"));
                Check(handler.Started.Wait(3000), "Ordinary tool is blocked in simulated Revit");
                var timer = Stopwatch.StartNew();
                using (var response = await client.PostAsync(endpoint, Body(RevitFailureStore.ToolName)))
                {
                    JObject json = JObject.Parse(await response.Content.ReadAsStringAsync());
                    Check(response.IsSuccessStatusCode && !(bool)json["result"]["isError"], "Cache tool succeeds over actual HTTP");
                    Check((int)json["result"]["structuredContent"]["total_count"] == 1, "HTTP captured result");
                }
                Check(timer.ElapsedMilliseconds < 3000 && !pending.IsCompleted, "Cache tool bypasses waiting ExternalEvent");
                using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint) { Content = Body(RevitFailureStore.ToolName) })
                {
                    request.Headers.Add("X-Revit-MCP-Client", "wpf_window");
                    using (var response = await client.SendAsync(request))
                        Check(response.IsSuccessStatusCode, "WPF-style HTTP call");
                }
                RevitMcpToolDefinition definition;
                Check(RevitMcpToolRegistry.TryGet(RevitFailureStore.ToolName, out definition) &&
                    definition.RiskLevel == RevitMcpToolRiskLevel.ReadOnly &&
                    (int)definition.InputSchema["properties"]["limit"]["default"] == 10, "Registry metadata");
            }
            finally
            {
                handler.Completion.TrySetResult(new RevitMcpToolCallResponse { Success = true, Result = new JObject() });
                if (pending != null) { using (await pending) { } }
                transport.Stop();
            }
        }
    }

    private static StringContent Body(string name)
    {
        return new StringContent(JsonConvert.SerializeObject(new JObject
        {
            { "jsonrpc", "2.0" }, { "id", 1 }, { "method", "tools/call" },
            { "params", new JObject { { "name", name }, { "arguments", new JObject() } } }
        }, Formatting.None), Encoding.UTF8, "application/json");
    }
}

// Replace only Revit execution and unrelated log/file services; use the production HTTP transport.
namespace KPLN_CoordiantorAI.ExternalAIModel.Mcp
{
    internal class RevitMcpExternalEventHandler
    {
        internal readonly ManualResetEventSlim Started = new ManualResetEventSlim();
        internal readonly TaskCompletionSource<RevitMcpToolCallResponse> Completion =
            new TaskCompletionSource<RevitMcpToolCallResponse>(TaskCreationOptions.RunContinuationsAsynchronously);
        public Task<RevitMcpToolCallResponse> ExecuteToolAsync(object doc, object ui, string name, JObject args, CancellationToken token)
        {
            Started.Set();
            return Completion.Task;
        }
    }
    internal static class RevitMcpDiagnosticLogger
    {
        public static void Log(string message) { }
        public static void LogException(string message, Exception error) { }
    }
    internal static class RevitMcpRequestCancellationRegistry
    {
        public static CancellationToken GetToken(string scope) { return CancellationToken.None; }
    }
    internal static class RevitMcpFileAttachmentService
    {
        public static bool IsFileTool(string name) { return false; }
        public static RevitMcpToolCallResponse ExecuteTool(string name, JObject args, CancellationToken token) { throw new NotSupportedException(); }
    }
    internal sealed class RevitMcpInteractionLogger
    {
        public RevitMcpInteractionLogger(string scenario) { }
        public void LogToolStart(string id, string name, JObject args, string area) { }
        public void LogToolError(string id, string name, JObject args, Exception ex, long ms, string area) { }
        public void LogToolEnd(string id, string name, JObject args, RevitMcpToolCallResponse result, long ms, string area) { }
    }
}
