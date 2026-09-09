using System;
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

        public MainThreadDispatcher()
        {
            // Form создаётся на текущем (главном) потоке — Execute() плагина
            // вызывается Navisworks именно там.
            _pump = new Form { ShowInTaskbar = false, WindowState = FormWindowState.Minimized };
            _pump.CreateControl(); // форсируем создание handle сразу
        }

        public T Run<T>(Func<T> func)
        {
            if (_pump.InvokeRequired)
                return (T)_pump.Invoke(func);
            return func();
        }

        public void Run(Action action)
        {
            if (_pump.InvokeRequired)
                _pump.Invoke(action);
            else
                action();
        }

        public void Dispose()
        {
            _pump.Invoke(new Action(() => _pump.Dispose()));
        }
    }
}
