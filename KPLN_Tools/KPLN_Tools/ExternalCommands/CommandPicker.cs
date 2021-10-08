using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using static KPLN_Loader.Output.Output;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandPicker: IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get application and documnet objects
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            //List of categories
            List<BuiltInCategory> defBuiltInCat = new List<BuiltInCategory>() { BuiltInCategory.OST_Grids, 
                                                                                BuiltInCategory.OST_Levels,
                                                                                BuiltInCategory.OST_ProjectBasePoint,
                                                                                BuiltInCategory.OST_RvtLinks 
                                                                                };

            //Get all element instances from list of builtInCat
            int cnt = 0;
            Transaction trans = new Transaction(doc);
            trans.Start("KPLN: Прикрепить элементы");
            foreach (BuiltInCategory curBuiltIn in defBuiltInCat)
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
