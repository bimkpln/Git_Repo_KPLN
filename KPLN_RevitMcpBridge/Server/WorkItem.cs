using System;
using System.Threading;
using System.Threading.Tasks;

namespace KPLN_RevitMcpBridge.Server
{
    // State transitions are atomic: a queued timeout can never execute later.
    internal sealed class WorkItem
    {
        private int _state; // 0 queued, 1 running, 2 completed, 3 cancelled
        public readonly string Id;
        public readonly string Payload;
        public readonly DateTime Created = DateTime.UtcNow;
        public readonly TaskCompletionSource<object> Completion = new TaskCompletionSource<object>(TaskCreationOptions.RunContinuationsAsynchronously);
        public WorkItem(string id, string payload) { Id = id; Payload = payload; }
        public string State => new[] { "queued", "running", "completed", "cancelled" }[Volatile.Read(ref _state)];
        public bool Cancel()
        {
            if (Interlocked.CompareExchange(ref _state, 3, 0) != 0) return false;
            Completion.TrySetResult(Json.Error(new BridgeException("cancelled", "Запрос отменён до начала выполнения.", 408)));
            return true;
        }
        public void Execute(Func<object> action)
        {
            if (Interlocked.CompareExchange(ref _state, 1, 0) != 0) return;
            object result;
            try { result = action(); }
            catch (Exception ex) { result = Json.Error(ex); }
            Completion.TrySetResult(result);
            Interlocked.Exchange(ref _state, 2);
        }
    }
}
