using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib.WorksetUtil;
using System;

namespace KPLN_ModelChecker_Debugger.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Worksetter : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Получение объектов приложения и документа
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            try
            {
                string folder = $"X:\\BIM\\5_Scripts\\Git_Repo_KPLN\\KPLN_ModelChecker_Debugger\\KPLN_ModelChecker_Debugger\\Workset_Patterns";
                if (WorksetSetService.ExecuteFromService(folder, doc))
                    return Result.Succeeded;
                else
                    return Result.Cancelled;
            }

            catch (Exception exc)
            {
                message = $"Произошла ошибка во время запуска скрипта - {exc}";
                return Result.Failed;
            }
        }
    }
}
