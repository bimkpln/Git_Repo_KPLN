﻿using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib;
using System;
using System.Collections.Generic;
using static KPLN_Loader.Output.Output;

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

            try
            {
                List<LevelAndGridSolid> sectDataSolids = LevelAndGridSolid.PrepareSolids(doc, "КП_О_Секция");

                using (Transaction t = new Transaction(doc))
                {
                    t.Start("КП: Построение боксов");

                    foreach (LevelAndGridSolid sectDataSolid in sectDataSolids)
                    {
                        DirectShape directShape = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
                        directShape.AppendShape(new GeometryObject[] { sectDataSolid.LevelSolid });
                        directShape.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set($"Секция: {sectDataSolid.GridData.CurrentSection}. Этаж: {sectDataSolid.CurrentLevelData.CurrentLevelNumber}");
                    }

                    t.Commit();
                }

            }
            catch (Exception ex)
            {
                if (ex is CheckerException _)
                {
                    TaskDialog taskDialog = new TaskDialog("ОШИБКА: Выполни инструкцию")
                    {
                        MainContent = ex.Message
                    };
                    taskDialog.Show();
                }

                else if (ex.InnerException != null)
                    Print($"Проверка не пройдена, работа скрипта остановлена. Передай ошибку: {ex.InnerException.Message}. StackTrace: {ex.StackTrace}", KPLN_Loader.Preferences.MessageType.Error);
                else
                    Print($"Проверка не пройдена, работа скрипта остановлена. Устрани ошибку: {ex.Message}. StackTrace: {ex.StackTrace}", KPLN_Loader.Preferences.MessageType.Error);

                return Result.Cancelled;
            }


            return Result.Succeeded;
        }
    }
}