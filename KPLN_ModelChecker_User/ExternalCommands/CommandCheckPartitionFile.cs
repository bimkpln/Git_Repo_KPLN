using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI;
using KPLN_ModelChecker_Lib;
using System;
using System.Collections.Generic;
using KPLN_ModelChecker_Lib.LevelAndGridBoxUtil;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandCheckPartitionFile : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // На будущее - может, имеет смысл посадить на конфиги, но в целом - лучше хардкодить,
            // чтобы никто случайно не влез. Конфиги нужно прятать от юзеров
            string sectParamName = "КП_О_Секция";
            string lvlIndexParamName = "КП_О_Этаж";
            if (doc.Title.StartsWith("СЕТ_1"))
            {
                sectParamName = "СМ_Секция";
                lvlIndexParamName = "СМ_Этаж";
            }

            try
            {
                List<LevelAndGridSolid> sectDataSolids = LevelAndGridSolid
                    .PrepareSolids(doc, sectParamName, lvlIndexParamName);

                using (Transaction t = new Transaction(doc))
                {
                    t.Start("KPLN: Построение боксов");

                    foreach (LevelAndGridSolid sectDataSolid in sectDataSolids)
                    {
                        DirectShape directShape = DirectShape
                            .CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
                        
                        directShape.AppendShape(new GeometryObject[] { sectDataSolid.CurrentSolid });
                        directShape
                            .get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
                            .Set($"Секция: {sectDataSolid.GridData.CurrentSection}. " +
                                 $"Этаж: {sectDataSolid.CurrentLevelData.CurrentLevelNumber}");
                    }

                    t.Commit();
                }

            }
            catch (Exception ex)
            {
                if (ex is CheckerException _)
                {
                    UserDialog ud = new UserDialog("ОШИБКА: Выполни инструкцию", ex.Message);
                    ud.Show();
                }

                else if (ex.InnerException != null)
                    Print(
                        $"Проверка не пройдена, работа скрипта остановлена. Передай ошибку: {ex.InnerException.Message}. StackTrace: {ex.StackTrace}",
                        MessageType.Error);
                else
                    Print(
                        $"Проверка не пройдена, работа скрипта остановлена. Устрани ошибку: {ex.Message}. StackTrace: {ex.StackTrace}",
                        MessageType.Error);

                return Result.Cancelled;
            }


            return Result.Succeeded;
        }
    }
}