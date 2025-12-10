using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Forms;


namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandNodeManager : IExternalCommand
    {
        internal const string PluginName = "Менеджер\nузлов";

        private static MainWindowNodeManager _window;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiapp = commandData.Application;
            var uidoc = uiapp.ActiveUIDocument;

            if (_window == null || !_window.IsLoaded)
            {
                _window = new MainWindowNodeManager(uiapp, uidoc);
                _window.Closed += (s, e) => _window = null; 
                _window.Show();
            }
            else
            {
                if (_window.WindowState == System.Windows.WindowState.Minimized)
                    _window.WindowState = System.Windows.WindowState.Normal;

                _window.Activate();
                _window.Topmost = true; 
                _window.Topmost = false;
            }

            return Result.Succeeded;
        }
    }
}
