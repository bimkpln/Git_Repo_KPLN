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
    public class Picker : IExternalCommand
    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Получение объектов приложения и документа
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            //Список используемых категорий
            List<BuiltInCategory> builtInCatList = new List<BuiltInCategory>() { 
                BuiltInCategory.OST_Grids,
                BuiltInCategory.OST_Levels,
                BuiltInCategory.OST_ProjectBasePoint,
                BuiltInCategory.OST_RvtLinks
            };

            //Элементы из списка категорий
            int cnt = 0;
            Transaction trans = new Transaction(doc);
            trans.Start("KPLN: Прикрепить элементы");
            foreach (BuiltInCategory curBuiltIn in builtInCatList)
            {
                IList<Element> elemColl = new FilteredElementCollector(doc)
                    .OfCategory(curBuiltIn)
                    .WhereElementIsNotElementType()
                    .ToElements();

                foreach (Element curElem in elemColl)
                {
                    if (!curElem.Pinned)
                    {
                        try
                        {
                            curElem.get_Parameter(BuiltInParameter.ELEMENT_LOCKED_PARAM).Set(1);
                            cnt++;
                        }
                        catch (Exception exc)
                        {
                            TaskDialog.Show("Ошибка: ", exc.ToString(), TaskDialogCommonButtons.Close);
                        }

                    }

                }
            }
            trans.Commit();
            if (cnt == 0)
            {
                TaskDialog.Show("Итог", "Все было прикреплено до запуска", TaskDialogCommonButtons.Ok);
            }
            else
            {
                string msg = "Прикреплено " + cnt + " элементов!";
                TaskDialog.Show("Итог", msg, TaskDialogCommonButtons.Ok);
            }
            return Result.Succeeded;
        }
    }
}
