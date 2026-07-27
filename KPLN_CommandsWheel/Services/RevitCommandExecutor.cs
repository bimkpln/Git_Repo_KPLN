using Autodesk.Revit.UI;
using KPLN_CommandsWheel.Models;
using System;

namespace KPLN_CommandsWheel.Services
{
    internal class RevitCommandExecutor
    {
        private readonly CommandRequestHandler _handler;
        private readonly ExternalEvent _externalEvent;

        internal RevitCommandExecutor()
        {
            _handler = new CommandRequestHandler();
            _externalEvent = ExternalEvent.Create(_handler);
        }

        internal void Run(RevitCommandInfo command)
        {
            if (command == null || string.IsNullOrWhiteSpace(command.Id))
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

        private class CommandRequestHandler : IExternalEventHandler
        {
            private readonly object _sync = new object();
            private string _commandId;
            private string _commandName;

            internal void SetCommand(string commandId, string commandName)
            {
                lock (_sync)
                {
                    _commandId = commandId;
                    _commandName = string.IsNullOrWhiteSpace(commandName) ? commandId : commandName;
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
                    revitCommandId = RevitCommandId.LookupCommandId(commandId);
                }
                catch
                {
                    revitCommandId = null;
                }

                if (revitCommandId == null)
                {
                    TaskDialog.Show("Команды", string.Format("Команда \"{0}\" не найдена в текущей сессии Revit.", commandName));
                    return;
                }

                bool canPost = false;
                try
                {
                    canPost = app.CanPostCommand(revitCommandId);
                }
                catch
                {
                    canPost = false;
                }

                if (!canPost)
                {
                    TaskDialog.Show("Команды", string.Format("Команда \"{0}\" сейчас недоступна на ленте Revit.", commandName));
                    return;
                }

                app.PostCommand(revitCommandId);
            }

            public string GetName()
            {
                return "KPLN Commands Wheel Command Runner";
            }
        }
    }
}