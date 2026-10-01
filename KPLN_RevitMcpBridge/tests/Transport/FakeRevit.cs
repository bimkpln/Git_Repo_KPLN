using System;
using System.Collections.Generic;
using System.Threading;
namespace Autodesk.Revit.UI
{
    public class UIApplication { }
    public interface IExternalEventHandler
    {
        void Execute(UIApplication app);
        string GetName();
    }
    public enum ExternalEventRequest { Accepted, Pending, Denied, TimedOut }
    public sealed class ExternalEvent : IDisposable
    {
        private readonly IExternalEventHandler _handler;
        private readonly int _uiThread = Thread.CurrentThread.ManagedThreadId;
        private readonly object _gate = new object();
        private bool _pending;
        public static ExternalEvent Latest;
        public static int PendingRaises, RaiseCalls, RaisesAfterDispose;
        public int RefuseRaises, ThrowRaises;
        public Action BeforeReturn;
        public bool Disposed { get; private set; }
        private ExternalEvent(IExternalEventHandler handler) { _handler = handler; }
        public static ExternalEvent Create(IExternalEventHandler handler) => Latest = new ExternalEvent(handler);
        public ExternalEventRequest Raise()
        {
            lock (_gate)
            {
                Interlocked.Increment(ref RaiseCalls);
                if (Disposed) { Interlocked.Increment(ref RaisesAfterDispose); throw new ObjectDisposedException("ExternalEvent"); }
                if (ThrowRaises > 0) { ThrowRaises--; throw new InvalidOperationException("transient signal failure"); }
                if (RefuseRaises > 0) { RefuseRaises--; return ExternalEventRequest.TimedOut; }
                if (_pending) { Interlocked.Increment(ref PendingRaises); return ExternalEventRequest.Pending; }
                _pending = true;
                return ExternalEventRequest.Accepted;
            }
        }
        // A background Revit window receives no simulated mouse/Idling events.
        // The test pumps this method only after the bridge wakes its window.
        public void Pump()
        {
            if (Thread.CurrentThread.ManagedThreadId != _uiThread) throw new Exception("Wrong event thread");
            lock (_gate) { if (!_pending || Disposed) return; }
            try { _handler.Execute(new UIApplication()); BeforeReturn?.Invoke(); }
            finally { lock (_gate) _pending = false; }
        }
        public void Dispose()
        {
            if (Thread.CurrentThread.ManagedThreadId != _uiThread) throw new Exception("Wrong dispose thread");
            lock (_gate) { Disposed = true; _pending = false; }
        }
    }
    public class ControlledApplication
    {
        public string VersionNumber => "2024";
        public event EventHandler DocumentChanged;
        public void Changed() { DocumentChanged?.Invoke(this, EventArgs.Empty); }
    }
    public class UIControlledApplication
    {
        public readonly ControlledApplication ControlledApplication = new ControlledApplication();
        private readonly int _uiThread = Thread.CurrentThread.ManagedThreadId;
        public IntPtr MainWindowHandle
        {
            get
            {
                if (Thread.CurrentThread.ManagedThreadId != _uiThread) throw new Exception("HWND read from worker thread");
                return new IntPtr(42);
            }
        }
    }
}
namespace KPLN_RevitMcpBridge.Services
{
    internal class RevitService
    {
        private readonly int _threadId = Thread.CurrentThread.ManagedThreadId;
        public static int Executed;
        public void DocumentChanged(object sender, EventArgs args) { }
        public object Execute(Autodesk.Revit.UI.UIApplication app, Dictionary<string, object> input)
        {
            if (Thread.CurrentThread.ManagedThreadId != _threadId) throw new Exception("Revit API called from HTTP thread");
            Interlocked.Increment(ref Executed);
            return new { ok = true, result = "executed" };
        }
    }
}
