using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using static KPLN_Loader.Output.Output;

namespace KPLN_Publication.ExternalCommands.BeforePublication
{
    [Transaction(TransactionMode.Manual)]
    class CommandRefreshSchedules : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                Document doc = commandData.Application.ActiveUIDocument.Document;

                List<ViewSheet> sheets = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>()
                    .Where(i => !i.IsPlaceholder)
                    .ToList();

                foreach (ViewSheet sheet in sheets)
                {
                    SchedulesRefresh.Start(doc, sheet);
                }
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
