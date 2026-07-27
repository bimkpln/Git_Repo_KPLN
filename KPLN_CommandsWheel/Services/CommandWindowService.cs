using Autodesk.Revit.UI;
using KPLN_CommandsWheel.Forms;
using KPLN_CommandsWheel.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_CommandsWheel.Services
{
    internal static class CommandWindowService
    {
        private static RevitCommandExecutor _executor;

        internal static bool ShowCommandSearch(UIApplication uiapp)
        {
            if (CommandSearchWindow.TryActivateExisting())
            {
                return true;
            }

            UserSettings settings = UserSettingsService.Load();
            List<RevitCommandInfo> commands = RibbonCommandCollector.Collect();
            SelectionCustomCommandService.AddCommands(commands);

            if (commands.Count == 0)
            {
                TaskDialog.Show("KPLN. Штурвал команд. Команды", "Не удалось прочитать команды ленты Revit в текущей сессии.");
                return false;
            }

            CommandSearchWindow window = new CommandSearchWindow(commands, settings, GetExecutor());
            WindowOwnerHelper.Apply(window);
            window.Show();

            return true;
        }

        internal static bool ShowCommandsWheel(UIApplication uiapp)
        {
            if (CommandsWheelWindow.TryActivateExisting())
            {
                return true;
            }

            UserSettings settings = UserSettingsService.Load();
            List<RevitCommandInfo> allCommands = RibbonCommandCollector.Collect();
            SelectionCustomCommandService.AddCommands(allCommands);

            Dictionary<string, RevitCommandInfo> commandsById = allCommands
                .GroupBy(command => command.Id, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);

            List<RevitCommandInfo> wheelCommands = new List<RevitCommandInfo>();
            foreach (string id in settings.WheelCommandIds.Take(8))
            {
                RevitCommandInfo command;
                if (!string.IsNullOrWhiteSpace(id) && commandsById.TryGetValue(id, out command))
                {
                    wheelCommands.Add(command);
                }
            }

            if (wheelCommands.Count == 0)
            {
                TaskDialog.Show("KPLN. Штурвал команд", "В штурвале нет доступных команд. Добавьте команды через окно \"Команды\".");
                return false;
            }

            CommandsWheelWindow window = new CommandsWheelWindow(wheelCommands, GetExecutor(), settings);
            WindowOwnerHelper.Apply(window);
            WindowPositionHelper.ShowCenteredOnCursor(window);
            window.Show();

            return true;
        }

        internal static void Shutdown()
        {
            RevitCommandExecutor executor = _executor;
            _executor = null;

            if (executor != null)
            {
                try
                {
                    executor.Dispose();
                }
                catch
                {
                    // Revit may already be releasing API resources during shutdown.
                }
                finally
                {
                    RibbonCommandCollector.ClearCache();
                }

                return;
            }

            RibbonCommandCollector.ClearCache();
        }

        private static RevitCommandExecutor GetExecutor()
        {
            if (_executor == null)
            {
                _executor = new RevitCommandExecutor();
            }

            return _executor;
        }
    }
}
