using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Forms;
using static KPLN_Loader.Output.Output;
using KPLN_Views_Ribbon.Forms;

namespace KPLN_Views_Ribbon.ExternalCommands.Views
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]

    class CommandBatchDelete : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                Document doc = commandData.Application.ActiveUIDocument.Document;

                List<ParameterFilterElement> filters = new FilteredElementCollector(doc)
                    .OfClass(typeof(ParameterFilterElement))
                    .Cast<ParameterFilterElement>()
                    .ToList();

                List<string> filterNames = filters.Select(x => x.Name).ToList();
                filterNames.Sort();

                FormBatchDelete form = new FormBatchDelete();
                form.Items = filterNames;

                form.ShowDialog();

                if (form.DialogResult != DialogResult.OK) return Result.Cancelled;

                List<string> deleteFilterNames = form.CheckedItems;

                List<ParameterFilterElement> filtersToDelete = filters
                    .Where(i => deleteFilterNames.Contains(i.Name))
                    .ToList();

                List<ElementId> ids = filtersToDelete.Select(i => i.Id).ToList();

                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Удаление фильтров: " + deleteFilterNames.Count.ToString());
                    doc.Delete(ids);
                    t.Commit();
                }

                form.Dispose();

                TaskDialog.Show("Удаление фильтров", "Успешно удалено фильтров: " + deleteFilterNames.Count.ToString());

                return Result.Succeeded;
            }

            catch (Exception e)
            {
                PrintError(e, "Произошла ошибка во время запуска скрипта");
                return Result.Failed;
            }

        }
    }
}
