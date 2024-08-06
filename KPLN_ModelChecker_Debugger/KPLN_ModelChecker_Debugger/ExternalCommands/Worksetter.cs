using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib.WorksetUtil;
using System;
using System.Collections.Generic;
using System.Linq;

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
                IEnumerable<RevitLinkInstance> links = new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>();

                if (WorksetSetService.ExecuteFromService(doc, links))
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
