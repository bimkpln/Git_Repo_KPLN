using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_CommandsWheel.Forms;
using KPLN_CommandsWheel.Models;
using KPLN_CommandsWheel.Services;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_CommandsWheel.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandsWheel : IExternalCommand
    {
        internal const string PluginName = "Комадна: Штурвал команд";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (CommandsWheelWindow.TryActivateExisting())
                return Result.Succeeded;

            UIApplication uiapp = commandData.Application;
            UserSettings settings = UserSettingsService.Load();
            List<RevitCommandInfo> allCommands = RibbonCommandCollector.Collect(uiapp);

            Dictionary<string, RevitCommandInfo> commandsById = allCommands
                .GroupBy(command => command.Id, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);

            List<RevitCommandInfo> wheelCommands = new List<RevitCommandInfo>();
            foreach (string id in settings.WheelCommandIds.Take(8))
            {
                if (!string.IsNullOrWhiteSpace(id) && commandsById.TryGetValue(id, out RevitCommandInfo command))
                    wheelCommands.Add(command);
            }

            if (wheelCommands.Count == 0)
            {
                TaskDialog.Show("KPLN. Штурвал команд", "В штурвале нет доступных команд. Добавьте команды через окно \"Команды\".");
                return Result.Cancelled;
            }

            RevitCommandExecutor executor = new RevitCommandExecutor();
            CommandsWheelWindow window = new CommandsWheelWindow(wheelCommands, executor);
            WindowOwnerHelper.Apply(window);
            window.Show();

            return Result.Succeeded;
        }
    }
}