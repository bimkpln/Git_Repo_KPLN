using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_TrailingMEP.Forms;
using KPLN_Library_Forms.Services;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_PluginActivityWorker;
using System;

namespace KPLN_TrailingMEP.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RouteTraceExtCmd : IExternalCommand
    {
        /// <summary>
        /// Имя плагина. Использую в KPLN_DefaultPanelExtension_Modify.
        /// </summary>
        public const string PluginName = "Дотянуть трассу";

        public Result ExecuteByUIApp(UIApplication uiapp)
        {
            RouteTraceForm mainForm = null;
            try
            {
                mainForm = new RouteTraceForm(uiapp);
                WindowHandleSearch.MainWindowHandle.SetAsOwner(mainForm);
                mainForm.Show();

                DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName(PluginName, ModuleData.ModuleName).ConfigureAwait(false);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                mainForm?.Close();
                HtmlOutput.PrintError(ex);
                return Result.Failed;
            }
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements) => ExecuteByUIApp(commandData.Application);
    }
}
