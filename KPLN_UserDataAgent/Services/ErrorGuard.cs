using Autodesk.Revit.UI;
using System;

namespace KPLN_UserDataAgent.Services
{
    internal sealed class ErrorGuard
    {
        private readonly object _lock = new object();
        private readonly bool _showDebugErrors;
        private DateTime _lastDialogTime = DateTime.MinValue;
        private string _pendingSource;
        private Exception _pendingException;

        public ErrorGuard(bool showDebugErrors)
        {
            _showDebugErrors = showDebugErrors;
        }

        public void Run(string source, Action action)
        {
            try
            {
                action();
            }
            catch (Exception exception)
            {
                HandleException(source, exception);
            }
        }

        public void HandleException(string source, Exception exception)
        {
            QueueException(source, exception);
            ShowPendingDialog();
        }

        public void QueueException(string source, Exception exception)
        {
            if (!_showDebugErrors)
                return;

            lock (_lock)
            {
                _pendingSource = source;
                _pendingException = exception;
            }
        }

        public void ShowPendingDialog()
        {
            if (!_showDebugErrors)
                return;

            DateTime now = DateTime.Now;
            if ((now - _lastDialogTime).TotalSeconds < ModuleData.DebugDialogMinIntervalSeconds)
                return;

            string source;
            Exception exception;
            lock (_lock)
            {
                source = _pendingSource;
                exception = _pendingException;
                _pendingSource = null;
                _pendingException = null;
            }

            if (exception == null)
                return;

            _lastDialogTime = now;

            try
            {
                TaskDialog.Show(
                    "KPLN UserDataAgent",
                    string.Format(
                        "Îøèáêà â {0}:{1}{2}",
                        source,
                        Environment.NewLine,
                        exception));
            }
            catch
            {
            }
        }
    }
}