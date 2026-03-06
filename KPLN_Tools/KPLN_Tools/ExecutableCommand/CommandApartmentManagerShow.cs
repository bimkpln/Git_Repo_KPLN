using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Common;
using System;
using System.Windows.Interop;

namespace KPLN_Tools.ExecutableCommand
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandApartmentManagerShow : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                var wnd = new KPLN_Tools.Forms.ApartmentManagerWindow(DBWorkerService.CurrentDBUserSubDepartment.Id);
                new WindowInteropHelper(wnd).Owner = uiapp.MainWindowHandle;

                bool? res = wnd.ShowDialog();
                if (res == true)
                {
                    int id = wnd.SelectedApartmentId;
                    // ТУТ: ВСЯ REVIT-ЛОГИКА ПО ID
                    TaskDialog.Show("Выбор", $"Выбран ID: {id}");
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка", ex.ToString());
                return Result.Failed;
            }
        }
    }
}