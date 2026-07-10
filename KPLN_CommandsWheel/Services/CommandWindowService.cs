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
        internal static bool ShowCommandSearch(UIApplication uiapp)
        {
            HotkeyService.Initialize();

            if (CommandSearchWindow.TryActivateExisting())
            {
                return true;
            }

            UserSettings settings = UserSettingsService.Load();
            List<RevitCommandInfo> commands = RibbonCommandCollector.Collect(uiapp);
            SelectionCustomCommandService.AddCommands(commands);

            if (commands.Count == 0)
            {
                TaskDialog.Show("KPLN. Штурвал команд. Команды", "Не удалось прочитать команды ленты Revit в текущей сессии.");
                return false;
            }

            RevitCommandExecutor executor = new RevitCommandExecutor();
            CommandSearchWindow window = new CommandSearchWindow(commands, settings, executor);
            WindowOwnerHelper.Apply(window);
            window.Show();

            return true;
        }

        internal static bool ShowCommandsWheel(UIApplication uiapp)
        {
            HotkeyService.Initialize();

            if (CommandsWheelWindow.TryActivateExisting())
            {
                return true;
            }

            UserSettings settings = UserSettingsService.Load();
            List<RevitCommandInfo> allCommands = RibbonCommandCollector.Collect(uiapp);
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

            RevitCommandExecutor executor = new RevitCommandExecutor();
            CommandsWheelWindow window = new CommandsWheelWindow(wheelCommands, executor, settings);
            WindowOwnerHelper.Apply(window);
            WindowPositionHelper.ShowCenteredOnCursor(window);
            window.Show();

            return true;
        }
    }
}