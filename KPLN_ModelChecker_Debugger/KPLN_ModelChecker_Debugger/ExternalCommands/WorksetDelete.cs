using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib.Forms;
using KPLN_ModelChecker_Lib.WorksetUtil;
using System.Linq;
using System.Windows;

namespace KPLN_ModelChecker_Debugger.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class WorksetDelete : IExternalCommand
    {
        internal const string PluginName = "Рабочие наборы - удалять";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Получение объектов приложения и документа
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            if (!doc.IsWorkshared)
            {
                MessageBox.Show(
                    "Работает только с моделями из хранилища",
                    PluginName,
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                return Result.Cancelled;
            }

            WorksetsForm emptyWorksetsForm = new WorksetsForm(doc, Util.GetDocWorksets(doc));
            emptyWorksetsForm.ShowDialog();

            return Result.Succeeded;
        }
    }
}
