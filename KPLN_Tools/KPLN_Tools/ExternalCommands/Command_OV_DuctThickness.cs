using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System.Linq;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class Command_OV_DuctThickness : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get application and documnet objects
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // НУЖНО ПРОСТО ПЕРЕНЕСТИ ИЗ PYREVIT С ОБЯЗАТЕЛЬНОЙ ФИКСАЦИЕЙ ЗАПУСКА ПЛАГИНА

            // Получаем используемые диаметры труб
            Pipe[] pipes = new FilteredElementCollector(doc)
                .OfClass(typeof(Pipe))
                .Cast<Pipe>()
                .ToArray();

            return Result.Cancelled;
        }
    }
}
