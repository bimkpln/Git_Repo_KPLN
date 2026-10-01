using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using KPLN_RevitMcpBridge.Server;
using KPLN_RevitMcpBridge.Services;

internal static class Program
{
    private static int _wakeMessages, _totalWakes;
    private static void Check(bool result, string message) { if (!result) throw new Exception(message); }
    private static string Body(Task<HttpResponseMessage> request) => request.GetAwaiter().GetResult().Content.ReadAsStringAsync().GetAwaiter().GetResult();
    private static Task<HttpResponseMessage> Post(HttpClient client, string body) => client.PostAsync("/command", new StringContent(body, Encoding.UTF8, "application/json"));
    private static string Payload(HttpBridgeServer bridge, string id = null) => Json.Serialize(new { session_id = bridge.SessionId, operation_id = id ?? Guid.NewGuid().ToString("N"), command = "test" });
    private static void Wake(IntPtr window)
    {
        Check(window == new IntPtr(42), "wrong background window");
        Interlocked.Increment(ref _wakeMessages);
        Interlocked.Increment(ref _totalWakes);
    }
    private static void PumpUntil(Func<bool> completed)
    {
        for (int i = 0; i < 300 && !completed(); i++)
        {
            if (Interlocked.Exchange(ref _wakeMessages, 0) != 0) ExternalEvent.Latest.Pump();
            Thread.Sleep(5);
        }
        Check(completed(), "background dispatch did not complete");
    }
    private static void WaitUntil(Func<bool> completed)
    {
        for (int i = 0; i < 100 && !completed(); i++) Thread.Sleep(5);
        Check(completed(), "request did not arrive");
    }
    private static void Main()
    {
        var directory = Path.Combine(Path.GetTempPath(), "KPLN-RevitMcpBridge-Test-" + Guid.NewGuid().ToString("N"));
        var app = new UIControlledApplication();
        using (var bridge = new HttpBridgeServer(app, directory, 1200, Wake))
        {
            bridge.Start();
            var discovery = Json.Parse(File.ReadAllText(Path.Combine(directory, bridge.SessionId + ".json")));
            using (var client = new HttpClient(new HttpClientHandler { UseProxy = false }) { BaseAddress = new Uri(bridge.Url), Timeout = TimeSpan.FromSeconds(5) })
            {
                Check(client.GetAsync("/health").Result.StatusCode == HttpStatusCode.Unauthorized, "unauthenticated health");
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + discovery.Text("token"));
                Check(Body(client.GetAsync("/health")).Contains(bridge.SessionId), "health identity");
                Check(Body(client.GetAsync("/health")).Contains("ExternalEvent"), "health dispatch context");
                client.DefaultRequestHeaders.Add("Origin", "http://example.com");
                Check(client.GetAsync("/health").Result.StatusCode == HttpStatusCode.Unauthorized, "browser origin accepted");
                client.DefaultRequestHeaders.Remove("Origin");
                var id = Guid.NewGuid().ToString("N");
                var payload = Json.Serialize(new { session_id = bridge.SessionId, operation_id = id, command = "test" });
                var response = Post(client, payload);
                PumpUntil(() => response.IsCompleted);
                Check(Body(response).Contains("executed") && RevitService.Executed == 1, "not dispatched to API thread");
                Check(Body(Post(client, payload)).Contains("executed") && RevitService.Executed == 1, "duplicate executed");
                Check(Post(client, payload.Replace("test", "other")).Result.StatusCode == HttpStatusCode.Conflict, "UUID payload collision");
                Check(Body(client.GetAsync("/operations/" + id)).Contains("completed"), "operation result absent");

                // Several queued requests must each get a UI callback with no
                // Idling subscription and no foreground/mouse event.
                var burst = new[] { Post(client, Payload(bridge)), Post(client, Payload(bridge)), Post(client, Payload(bridge)) };
                PumpUntil(() => Array.TrueForAll(burst, p => p.IsCompleted));
                foreach (var request in burst) Check(Body(request).Contains("executed"), "burst request stranded");
                Check(RevitService.Executed == 4, "burst execution count");

                ExternalEvent.Latest.RefuseRaises = 1;
                var refused = Post(client, Payload(bridge));
                PumpUntil(() => refused.IsCompleted);
                Check(Body(refused).Contains("executed"), "TimedOut signal was not retried");
                ExternalEvent.Latest.ThrowRaises = 1;
                var failedSignal = Post(client, Payload(bridge));
                PumpUntil(() => failedSignal.IsCompleted);
                Check(Body(failedSignal).Contains("executed"), "signal exception stranded request");

                // A request arrives after Execute drains the queue, while Revit
                // still considers the previous event pending until return.
                Task<HttpResponseMessage> boundary = null;
                ExternalEvent.Latest.BeforeReturn = () =>
                {
                    ExternalEvent.Latest.BeforeReturn = null;
                    int pendingBefore = Volatile.Read(ref ExternalEvent.PendingRaises);
                    boundary = Post(client, Payload(bridge));
                    WaitUntil(() => Volatile.Read(ref ExternalEvent.PendingRaises) > pendingBefore);
                };
                var beforeBoundary = Post(client, Payload(bridge));
                PumpUntil(() => beforeBoundary.IsCompleted && boundary != null && boundary.IsCompleted);
                Check(Body(beforeBoundary).Contains("executed") && Body(boundary).Contains("executed"), "event return race lost request");
                Check(RevitService.Executed == 8, "requests executed more than once");

                var timedId = Guid.NewGuid().ToString("N");
                var timed = Body(Post(client, Json.Serialize(new { session_id = bridge.SessionId, operation_id = timedId, command = "test" })));
                Check(timed.Contains("queue_timeout"), "queued timeout missing");
                ExternalEvent.Latest.Pump();
                Check(RevitService.Executed == 8, "timed out command executed later");
                Check(Body(client.GetAsync("/operations/" + timedId)).Contains("cancelled"), "cancel status absent");
                Thread.Sleep(300);
                int wakesAfterDrain = Volatile.Read(ref _totalWakes);
                Thread.Sleep(300);
                Check(Volatile.Read(ref _totalWakes) == wakesAfterDrain, "idle bridge keeps waking Revit");
                Check(Post(client, "{bad}").Result.StatusCode == HttpStatusCode.BadRequest, "invalid JSON accepted");
                Check(Post(client, Json.Serialize(new { session_id = "wrong", operation_id = Guid.NewGuid().ToString("N") })).Result.StatusCode == HttpStatusCode.Conflict, "wrong session accepted");

                int raisesBeforeStop = Volatile.Read(ref ExternalEvent.RaiseCalls);
                var duringStop = Post(client, Payload(bridge));
                WaitUntil(() => Volatile.Read(ref ExternalEvent.RaiseCalls) > raisesBeforeStop);
                var stoppedEvent = ExternalEvent.Latest;
                bridge.Dispose();
                Thread.Sleep(300);
                stoppedEvent.Pump();
                Check(stoppedEvent.Disposed && ExternalEvent.RaisesAfterDispose == 0, "event/timer shutdown race");
                Check(RevitService.Executed == 8, "stopped request executed");
                try { Body(duringStop); } catch (HttpRequestException) { } // listener can close the response
            }
        }
        Check(Directory.GetFiles(directory, "*.json").Length == 0, "discovery not removed on shutdown");
        using (var restarted = new HttpBridgeServer(app, directory, 1200, Wake))
        {
            restarted.Start();
            var discovery = Json.Parse(File.ReadAllText(Path.Combine(directory, restarted.SessionId + ".json")));
            using (var client = new HttpClient(new HttpClientHandler { UseProxy = false }) { BaseAddress = new Uri(restarted.Url) })
            {
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + discovery.Text("token"));
                var response = Post(client, Payload(restarted));
                PumpUntil(() => response.IsCompleted);
                Check(Body(response).Contains("executed") && RevitService.Executed == 9, "restart did not dispatch cleanly");
            }
        }
        Directory.Delete(directory); // exact generated test directory, nonrecursive and empty
        Console.WriteLine("PASS: loopback HTTP, auth, Origin, background ExternalEvent + wakeup, burst, retry, pending return race, API thread, deduplication, busy timeout cancellation, idle silence, stop/restart (Revit API doubles).");
    }
}
