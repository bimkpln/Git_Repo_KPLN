using System;
using System.Threading;
using System.Threading.Tasks;
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
        Console.WriteLine("PASS: cancel before execution, single execution, exception result, running timeout, 2000 cancellation races.");
    }
}
