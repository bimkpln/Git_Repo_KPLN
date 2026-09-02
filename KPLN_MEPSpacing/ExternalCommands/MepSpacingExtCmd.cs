using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_MEPSpacing.Forms;
using KPLN_Library_Forms.Services;
using KPLN_Library_Forms.UI.HtmlWindow;
using System;
using System.Collections.Generic;

namespace KPLN_MEPSpacing.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MepSpacingExtCmd : IExternalCommand
    {
        /// <summary>
        /// Имя плагина. Использую в KPLN_DefaultPanelExtension_Modify.
        /// </summary>
        public const string PluginName = "Выровнять шаг";

        public Result ExecuteByUIApp(UIApplication uiapp)
        {
            try
            {
                UIDocument uiDoc = uiapp.ActiveUIDocument;
                if (uiDoc == null)
                    return Result.Cancelled;

                MepSpacingForm mainForm = new MepSpacingForm(uiapp, new List<ElementId>());
                WindowHandleSearch.MainWindowHandle.SetAsOwner(mainForm);
                mainForm.Show();

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                HtmlOutput.PrintError(ex);
                return Result.Failed;
            }
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements) =>
            ExecuteByUIApp(commandData.Application);
    }
}
