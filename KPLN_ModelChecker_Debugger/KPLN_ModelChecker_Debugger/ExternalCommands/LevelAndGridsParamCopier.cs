using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_ModelChecker_Debugger.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal sealed class LevelAndGridsParamCopier : IExternalCommand
    {
        private IReadOnlyCollection<BuiltInCategory> _gridAndLevelsCatList = new List<BuiltInCategory>() {
                BuiltInCategory.OST_Grids,
                BuiltInCategory.OST_Levels,
            };

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Получение объектов приложения и документа
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            
            // Получаю стеки из проекта
            IEnumerable<Element> elemColl = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Grids).WhereElementIsNotElementType();
            elemColl = elemColl.Concat(new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType());

            
            

            //Список используемых категорий
            List<BuiltInCategory> builtInCatList = new List<BuiltInCategory>() {
                BuiltInCategory.OST_Grids,
                BuiltInCategory.OST_Levels,
                BuiltInCategory.OST_ProjectBasePoint,
                BuiltInCategory.OST_RvtLinks
            };

            Transaction trans = new Transaction(doc);
            trans.Start("KPLN: Параметры сеток");
            trans.Commit();

            return Result.Succeeded;
        }
    }
}
