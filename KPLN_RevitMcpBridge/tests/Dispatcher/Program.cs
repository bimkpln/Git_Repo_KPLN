using System;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using KPLN_RevitMcpBridge.Server;

internal static class Program
{
    private static void Assert(bool condition, string message) { if (!condition) throw new Exception(message); }
    private static void Main()
    {
        int count = 0;
        var cancelled = new WorkItem("cancel", "{}");
        Assert(cancelled.Cancel(), "cancel queued");
        cancelled.Execute(() => { count++; return null; });
        Assert(count == 0 && cancelled.State == "cancelled", "cancelled write executed");
        var complete = new WorkItem("complete", "{}");
        complete.Execute(() => { count++; return 42; });
        complete.Execute(() => { count++; return 99; });
        Assert(count == 1 && (int)complete.Completion.Task.Result == 42, "duplicate execution");
        var failure = new WorkItem("failure", "{}");
        failure.Execute(() => { throw new InvalidOperationException("test error"); });
        Assert(Json.Serialize(failure.Completion.Task.Result).Contains("test error"), "exception missing");
        using (var started = new ManualResetEventSlim())
        using (var release = new ManualResetEventSlim())
        {
            var running = new WorkItem("running", "{}");
            var worker = Task.Run(() => running.Execute(() => { started.Set(); release.Wait(); return 7; }));
            started.Wait(); Assert(!running.Cancel(), "running request reported cancelled");
            release.Set(); worker.Wait(); Assert((int)running.Completion.Task.Result == 7, "running result lost");
        }
        for (int i = 0; i < 2000; i++)
        {
            int writes = 0; bool didCancel = false;
            var item = new WorkItem(i.ToString(), "{}");
            Parallel.Invoke(() => didCancel = item.Cancel(), () => item.Execute(() => { Interlocked.Increment(ref writes); return null; }));
            Assert(writes <= 1 && (!didCancel || writes == 0) && item.Completion.Task.IsCompleted, "cancel/execute race");
        }
        var history = new Dictionary<string, WorkItem>();
        var protectedWrite = new WorkItem("write", "{}");
        protectedWrite.Execute(() => 123);
        history.Add(protectedWrite.Id, protectedWrite);
        var pendingRead = new WorkItem("pending", "{}", true);
        history.Add(pendingRead.Id, pendingRead);
        for (int i = 0; i < 2000; i++)
        {
            WorkItem.TrimCompletedReads(history, 512);
            Assert(history.Count < 512, "completed reads blocked next request");
            var read = new WorkItem("read" + i, "{}", true);
            read.Execute(() => i);
            history.Add(read.Id, read);
        }
        Assert(history.Count == 512, "history is not bounded");
        Assert(history.ContainsKey("write") && history.ContainsKey("pending"), "write or pending read evicted");
        Assert((int)history["write"].Completion.Task.Result == 123, "write receipt lost");
        Assert(WorkItem.IsReadOnlyCommand("get_view_annotations"), "annotation query not read-only");
        foreach (var command in new[] { "create_tags", "set_parameters", "export_selected_sheets_pdf", "print_selected_sheets_pdf", "get_unknown", "" })
            Assert(!WorkItem.IsReadOnlyCommand(command), "unsafe command eviction: " + command);
        var writesOnly = new Dictionary<string, WorkItem>();
        for (int i = 0; i < 512; i++)
        {
            var write = new WorkItem(i.ToString(), "{}");
            write.Execute(() => 1);
            writesOnly.Add(write.Id, write);
        }
        WorkItem.TrimCompletedReads(writesOnly, 512);
        Assert(writesOnly.Count == 512, "protected write receipts evicted");
        Console.WriteLine("PASS: dispatcher races, 2000 read requests, bounded history, protected writes and pending reads.");
    }
}
