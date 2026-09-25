using System;
using System.Threading;
using System.Windows.Forms;

namespace KPLN_NavisMcpBridge
{
    /// <summary>
    /// Navisworks API можно дёргать только из главного потока приложения.
    /// HttpListener присылает запросы на своих собственных потоках, поэтому
    /// каждый вызов в ClashService должен идти через этот диспетчер
    /// (скрытая форма + Control.Invoke — стандартный трюк для маршалинга
    /// в UI-поток из фонового потока в WinForms-хосте, каким является
    /// процесс Navisworks).
    /// </summary>
    internal sealed class MainThreadDispatcher : IDisposable
    {
        private readonly Form _pump;
        private readonly int _ownerThreadId;

        public MainThreadDispatcher()
        {
            // Form создаётся на текущем (главном) потоке — Execute() плагина
            // вызывается Navisworks именно там.
            _ownerThreadId = Thread.CurrentThread.ManagedThreadId;
            _pump = new Form { ShowInTaskbar = false, WindowState = FormWindowState.Minimized };

            // CreateControl() не гарантирует создание handle у невидимой
            // формы. Без handle InvokeRequired на фоновом HTTP-потоке может
            // вернуть false, после чего Navisworks API вызывается не из UI-
            // потока и запрос зависает. Обращение к Handle создаёт его сразу
            // на потоке-владельце.
            _ = _pump.Handle;
        }

        public T Run<T>(Func<T> func)
        {
            ThrowIfDisposed();
            if (Thread.CurrentThread.ManagedThreadId != _ownerThreadId)
                return (T)_pump.Invoke(func);
            return func();
        }

        public void Run(Action action)
        {
            ThrowIfDisposed();
            if (Thread.CurrentThread.ManagedThreadId != _ownerThreadId)
                _pump.Invoke(action);
            else
                action();
        }

        public void Dispose()
        {
            if (_pump.IsDisposed)
                return;

            if (Thread.CurrentThread.ManagedThreadId != _ownerThreadId)
                _pump.Invoke(new Action(() => _pump.Dispose()));
            else
                _pump.Dispose();
        }

        private void ThrowIfDisposed()
        {
            if (_pump.IsDisposed || !_pump.IsHandleCreated)
                throw new ObjectDisposedException(nameof(MainThreadDispatcher));
        }
    }
}
