using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace KPLN_TaskManager.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ShowMainMenu : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            ShowPanel(commandData.Application, true);

            return Result.Succeeded;
        }

        internal static void ShowPanel(object app, bool isForceStart)
        {
            DockablePane dockableWindow;
            if (app is UIControlledApplication uiCA)
                dockableWindow = uiCA.GetDockablePane(Module.PaneId);
            else if (app is UIApplication uiapp)
                dockableWindow = uiapp.GetDockablePane(Module.PaneId);
            else
                throw new System.Exception("Ошибка приведения типа. Обратись к разработчику!!!");

            if (isForceStart)
                dockableWindow.Show();
            else if (Module.MainMenuViewer.FilteredTasks != null && !Module.MainMenuViewer.FilteredTasks.IsEmpty)
                dockableWindow.Show();
            else
                dockableWindow.Hide();
        }
    }
}
