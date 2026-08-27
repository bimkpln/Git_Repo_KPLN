using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_MEPBender.Forms;
using System;
using System.Windows.Interop;

namespace KPLN_MEPBender.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MepBenderExtCmd : IExternalCommand
    {
        public const string PluginName = "Обход пересечений";

        public Result ExecuteByUIApp(UIApplication uiapp)
        {
            try
            {
                if (MepBenderForm.TryActivateExisting())
                    return Result.Succeeded;

                MepBenderForm mainForm = new MepBenderForm(uiapp);
                new WindowInteropHelper(mainForm).Owner = uiapp.MainWindowHandle;
                mainForm.Show();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show(PluginName, ex.ToString());
                return Result.Failed;
            }
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements) => ExecuteByUIApp(commandData.Application);
    }
}
