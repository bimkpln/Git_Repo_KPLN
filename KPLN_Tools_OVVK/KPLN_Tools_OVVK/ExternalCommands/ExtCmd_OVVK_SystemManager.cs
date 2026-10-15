using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.Services;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Tools_OVVK.Forms;
using System;
using System.Windows;

namespace KPLN_Tools_OVVK.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal sealed class ExtCmd_OVVK_SystemManager : IExternalCommand
    {
        internal const string PluginName = "Менеджер систем";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return ExecuteByUIApp(commandData.Application);
        }


        internal Result ExecuteByUIApp(UIApplication app)
        {
            OVVK_SystemManagerForm form = null;

            try
            {
                if (OVVK_SystemManagerForm.TryActivateExisting())
                    return Result.Succeeded;

                Document doc = app.ActiveUIDocument?.Document;
                if (doc == null || doc.IsFamilyDocument)
                    return Result.Cancelled;


                form = new OVVK_SystemManagerForm(doc);
                new WindowHandleSearch(app.MainWindowHandle).SetAsOwner(form);
                if (form.CurrentOVVK_SystemManagerVM.CurrentOVVK_SystemManagerM.ElementColl.Count == 0)
                {
                    MessageBox.Show("В модели отсутствуют элементы ОВВК для анализа.", "KPLN: Внимание", MessageBoxButton.OK);
                    form.Close();
                    return Result.Cancelled;
                }

                form.StartTracking(app);
                form.Show();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                form?.Close();
                HtmlOutput.PrintError(ex);
                return Result.Failed;
            }
        }
    }
}
