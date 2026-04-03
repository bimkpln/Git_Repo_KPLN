using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ViewsAndLists_Ribbon.Forms;

namespace KPLN_ViewsAndLists_Ribbon.ExternalCommands.Lists
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class ExtCmdListIZMFilling : IExternalCommand
    {
        internal const string PluginName = "Заполнение ИЗМов";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            FrmIZMFilling wnd = new FrmIZMFilling(doc);
            wnd.ShowDialog();

            return Result.Succeeded;
        }
    }
}