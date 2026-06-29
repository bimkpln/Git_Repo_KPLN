using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_CommandsWheel.Forms;
using KPLN_CommandsWheel.Models;
using KPLN_CommandsWheel.Services;
using System.Collections.Generic;

namespace KPLN_CommandsWheel.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandSearch : IExternalCommand
    {
        internal const string PluginName = "Окно команд";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UserSettings settings = UserSettingsService.Load();
            List<RevitCommandInfo> commands = RibbonCommandCollector.Collect(uiapp);

            if (commands.Count == 0)
            {
                TaskDialog.Show("KPLN. Штурвал команд. Команды", "Не удалось прочитать команды ленты Revit в текущей сессии.");
                return Result.Cancelled;
            }

            RevitCommandExecutor executor = new RevitCommandExecutor();
            CommandSearchWindow window = new CommandSearchWindow(commands, settings, executor);
            WindowOwnerHelper.Apply(window);
            window.Show();

            return Result.Succeeded;
        }
    }
}