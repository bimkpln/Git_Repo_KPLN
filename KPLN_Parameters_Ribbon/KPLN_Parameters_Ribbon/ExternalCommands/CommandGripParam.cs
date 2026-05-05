using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_DBWorker;
using KPLN_Parameters_Ribbon.Common.GripParam;
using KPLN_Parameters_Ribbon.Common.GripParam.Builder;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Parameters_Ribbon.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandGripParam : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Получение объектов приложения и документа
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            AbstrGripBuilder gripBuilder = null;
            int userDepartment = SQLiteMainService.CurrentDBUser.SubDepartmentId;
            try
            {
                string docPath = doc.Title.ToUpper();
                // Посадить на конфиг под каждый файл
                if (userDepartment == 2 || userDepartment == 8 && docPath.Contains("_АР"))
                {
                    if (docPath.StartsWith("ОБДН_"))
                        gripBuilder = new GripBuilder_AR(doc, "ОБДН", "SMNX_Этаж", "SMNX_Секция");

                    else if (docPath.StartsWith("ИЗМЛ_"))
                        gripBuilder = new GripBuilder_AR(doc, "ИЗМЛ", "КП_О_Этаж", "КП_О_Секция");

                    else if (docPath.StartsWith("СЕТ_1_"))
                        gripBuilder = new GripBuilder_AR(doc, "СЕТ_1", "СМ_Этаж", "СМ_Секция");

                    else if (docPath.StartsWith("ОМК3_"))
                        gripBuilder = new GripBuilder_AR(doc, "ОМК3", "КП_О_Этаж", "КП_О_Секция");

                    else if (docPath.StartsWith("СГРВН_"))
                        gripBuilder = new GripBuilder_AR(doc, "СГРВН", "КП_О_Этаж", "КП_О_Секция", "КП_О_Корпус");
                }
                else if (userDepartment == 3 || userDepartment == 8 && docPath.Contains("_КР"))
                {
                    if (docPath.StartsWith("ИЗМЛ_"))
                        gripBuilder = new GripBuilder_KR(doc, "ИЗМЛ", "О_Этаж", "КП_О_Секция");

                    else if (docPath.StartsWith("СЕТ_1_"))
                        gripBuilder = new GripBuilder_KR(doc, "СЕТ_1", "СМ_Этаж", "СМ_Секция");

                    else if (docPath.StartsWith("ОМК3_"))
                        gripBuilder = new GripBuilder_KR(doc, "ОМК3", "О_Этаж", "КП_О_Секция");

                    else if (docPath.StartsWith("СГРВН_"))
                        gripBuilder = new GripBuilder_KR(doc, "СГРВН", "О_Этаж", "КП_О_Секция", "КП_О_Корпус");
                }
                else if (userDepartment == 4
                         || userDepartment == 5
                         || userDepartment == 6
                         || userDepartment == 7
                         || userDepartment == 8
                         && (docPath.Contains("_ОВ")
                             || docPath.Contains("_ВК")
                             || docPath.Contains("_АУПТ")
                             || docPath.Contains("_ПТ")
                             || docPath.Contains("_ЭОМ")
                             || docPath.Contains("_СС")
                             || docPath.Contains("_ПБ")
                             || docPath.Contains("_АК")
                             || docPath.Contains("_АВ")))
                {
                    if (docPath.StartsWith("ОБДН_"))
                        gripBuilder = new GripBuilder_IOS(doc, "ОБДН_", "SMNX_Этаж", "SMNX_Секция");

                    else if (docPath.StartsWith("ИЗМЛ_"))
                        gripBuilder = new GripBuilder_IOS(doc, "ИЗМЛ", "КП_О_Этаж", "КП_О_Секция");

                    else if (docPath.StartsWith("ПШМ1_"))
                        gripBuilder = new GripBuilder_IOS(doc, "ИЗМЛ", "КП_О_Этаж", "КП_О_Секция");

                    else if (docPath.StartsWith("СЕТ_1_"))
                        gripBuilder = new GripBuilder_IOS(doc, "СЕТ_1", "СМ_Этаж", "СМ_Секция");

                    else if (docPath.StartsWith("ОМК3_"))
                        gripBuilder = new GripBuilder_IOS(doc, "ОМК3", "КП_О_Этаж", "КП_О_Секция");

                    else if (docPath.StartsWith("СГРВН_"))
                        gripBuilder = new GripBuilder_IOS(doc, "СГРВН", "КП_О_Этаж", "КП_О_Секция", "КП_О_Корпус");
                }
                else
                {
                    TaskDialog taskDialog = new TaskDialog("ОШИБКА: Выполни инструкцию")
                    {
                        MainContent = "Ошибка определения проекта/пользователя. Обратись в BIM-отдел",
                        MainIcon = TaskDialogIcon.TaskDialogIconInformation
                    };
                    taskDialog.Show();
                    return Result.Failed;
                }

                GripDirector gripDirector = new GripDirector(gripBuilder);
                gripDirector.BuildWriter();
                if (gripBuilder.ErrorElements.Count > 0)
                {
                    HashSet<string> uniqErrors = new HashSet<string>(gripBuilder.ErrorElements.Select(e => e.ErrorMessage));
                    foreach (string error in uniqErrors)
                    {
                        List<string> errorElements = gripBuilder
                            .ErrorElements
                            .Where(e => e.ErrorMessage.Equals(error))
                            .Select(e => e.ErrorElement.Id.ToString())
                            .ToList();

                        StringBuilder errorIdCollBuilder = new StringBuilder();
                        for (int i = 0; i < errorElements.Count; i++)
                        {
                            if (i > 0)
                                errorIdCollBuilder.Append(",");
                            if (i > 0 && i % 12 == 0)
                                errorIdCollBuilder.AppendLine();

                            errorIdCollBuilder.Append(errorElements[i]);
                        }

                        Print($"{error} - для след. элементов:\n {errorIdCollBuilder}",
                            MessageType.Warning);
                    }
                }

                return Result.Succeeded;
            }
            catch (GripParamExection gpe)
            {
                string msg = string.Empty;
                if (gpe.ErrorElements == null)
                    msg = $"ОШИБКА:\n{gpe.ErrorMessage}";
                else
                    msg = $"ОШИБКА:\n{gpe.ErrorMessage}\nДля элементов:\n{string.Join(", ", gpe.ErrorElements.Select(e => e.Id.ToString()))}";

                TaskDialog taskDialog = new TaskDialog("ОШИБКА: Выполни инструкцию")
                {
                    MainContent = msg,
                    MainIcon = TaskDialogIcon.TaskDialogIconWarning
                };
                taskDialog.Show();

                return Result.Failed;
            }
            // Обработка для многопоточки (т.е. несколько ошибок, которые могут произойти при вызове функций несколько раз в разных потоках)
            catch (AggregateException ae)
            {
                HashSet<Exception> hashAeS = ae.InnerExceptions.ToHashSet();
                foreach (Exception e in hashAeS)
                {
                    if (e is GripParamExection gpe)
                        Print($"Прервано с ошибкой: {gpe.ErrorMessage}", MessageType.Error);
                    else
                        Print($"Прервано с ошибкой: {e}", MessageType.Error);
                }

                return Result.Failed;
            }
            catch (Exception e)
            {
                Print($"Прервано с ошибкой: {e}", MessageType.Error);

                return Result.Failed;
            }
        }
    }
}