using Autodesk.Revit.UI;
using ComponentManager = Autodesk.Windows.ComponentManager;
using KPLN_CommandsWheel.Models;
using System;
using System.Windows.Threading;

namespace KPLN_CommandsWheel.Services
{
    internal sealed class RevitCommandExecutor : IDisposable
    {
        private readonly CommandRequestHandler _handler;
        private readonly ExternalEvent _externalEvent;
        private bool _isDisposed;

        internal RevitCommandExecutor()
        {
            _handler = new CommandRequestHandler();
            _externalEvent = ExternalEvent.Create(_handler);
        }

        internal void Run(RevitCommandInfo command)
        {
            if (_isDisposed || command == null || string.IsNullOrWhiteSpace(command.Id))
            {
                return;
            }

            _handler.SetCommand(command.Id, command.Name);

            try
            {
                _externalEvent.Raise();
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Команды", "Не удалось передать команду в Revit:\n" + ex.Message);
            }
        }

        public void Dispose()
        {
            if (_isDisposed)
            {
                return;
            }

            _isDisposed = true;
            _handler.Clear();
            _externalEvent.Dispose();
        }

        private class CommandRequestHandler : IExternalEventHandler
        {
            private readonly object _sync = new object();
            private string _commandId;
            private string _commandName;
            private bool _stopped;
            private bool _ribbonCommandQueued;

            internal void SetCommand(string commandId, string commandName)
            {
                lock (_sync)
                {
                    if (_stopped)
                    {
                        return;
                    }

                    _commandId = commandId;
                    _commandName = string.IsNullOrWhiteSpace(commandName) ? commandId : commandName;
                }
            }

            internal void Clear()
            {
                lock (_sync)
                {
                    _stopped = true;
                    _commandId = null;
                    _commandName = null;
                }
            }

            public void Execute(UIApplication app)
            {
                string commandId;
                string commandName;

                lock (_sync)
                {
                    commandId = _commandId;
                    commandName = _commandName;
                    _commandId = null;
                    _commandName = null;
                }

                if (string.IsNullOrWhiteSpace(commandId))
                {
                    return;
                }

                if (SelectionCustomCommandService.TryExecute(app, commandId))
                {
                    return;
                }

                RevitCommandId revitCommandId = null;
                try
                {
                    revitCommandId = RibbonCommandCollector.ResolveCommandId(commandId);
                }
                catch
                {
                    revitCommandId = null;
                }

                bool canPost = false;
                try
                {
                    canPost = revitCommandId != null && app.CanPostCommand(revitCommandId);
                }
                catch
                {
                    canPost = false;
                }

                if (!canPost)
                {
                    QueueRibbonCommand(commandId, commandName);
                    return;
                }

                if (!RibbonCommandCollector.IsAvailableInCurrentContext(commandId))
                {
                    ShowUnavailable(commandName);
                    return;
                }

                try
                {
                    app.PostCommand(revitCommandId);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        "KPLN Commands Wheel: cannot post command. " + ex.Message);
                    ShowUnavailable(commandName);
                }
            }

            private void QueueRibbonCommand(string commandId, string commandName)
            {
                Dispatcher dispatcher = ComponentManager.Ribbon == null
                    ? null : ComponentManager.Ribbon.Dispatcher;
                if (_stopped || _ribbonCommandQueued || dispatcher == null
                    || dispatcher.HasShutdownStarted || dispatcher.HasShutdownFinished)
                {
                    ShowUnavailable(commandName);
                    return;
                }

                _ribbonCommandQueued = true;
                try
                {
                    // Native ribbon handlers must run like a normal UI click,
                    // after the API callback and the wheel's mouse event return.
                    dispatcher.BeginInvoke(new Action(delegate
                    {
                        try
                        {
                            if (!_stopped && !RibbonCommandCollector.TryExecuteRibbonCommand(commandId))
                            {
                                ShowUnavailable(commandName);
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine(
                                "KPLN Commands Wheel: ribbon dispatch failed. " + ex.Message);
                            ShowUnavailable(commandName);
                        }
                        finally
                        {
                            _ribbonCommandQueued = false;
                        }
                    }), DispatcherPriority.ApplicationIdle);
                }
                catch (Exception ex)
                {
                    _ribbonCommandQueued = false;
                    System.Diagnostics.Debug.WriteLine(
                        "KPLN Commands Wheel: cannot queue ribbon command. " + ex.Message);
                    ShowUnavailable(commandName);
                }
            }

            private static void ShowUnavailable(string commandName)
            {
                TaskDialog.Show("Команды", string.Format(
                    "В данном контексте команду «{0}» выполнить нельзя.", commandName));
            }

            public string GetName()
            {
                return "KPLN Commands Wheel Command Runner";
            }
        }
    }
}
