using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Parameters_Ribbon.Common.GripParam;
using KPLN_Parameters_Ribbon.Common.GripParam.Builder;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Loader.Output.Output;

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
            int userDepartment = KPLN_Loader.Preferences.User.Department.Id;
            // Техническая подмена оазделов для режима тестирования
            if (userDepartment == 6) { userDepartment = 4; }
            
            try
            {
                string docPath = doc.Title.ToUpper();
                // Посадить на конфиг под каждый файл
                if (docPath.Contains("АР"))
                {
                    if (docPath.Contains("ОБДН"))
                    {
                        gripBuilder = new GripBuilder_AR(doc, "ОБДН", "SMNX_Этаж", 1, "SMNX_Секция", 0.328, 3);
                    }
                    else if (docPath.Contains("ИЗМЛ"))
                    {
                        gripBuilder = new GripBuilder_AR(doc, "ИЗМЛ", "КП_О_Этаж", 1, "КП_О_Секция", 0.328, 10);
                    }
                }
                else if (docPath.Contains("КР"))
                {
                    if (docPath.Contains("ОБДН"))
                    {
                        gripBuilder = new GripBuilder_AR(doc, "ОБДН", "О_Этаж", 1, "SMNX_Секция", 0.328, 3);
                    }
                    else if (docPath.Contains("ИЗМЛ"))
                    {
                        gripBuilder = new GripBuilder_KR(doc, "ИЗМЛ", "О_Этаж", 1, "КП_О_Секция", 0.328, 10);
                    }
                }
                else if ((docPath.Contains("ОВ") || docPath.Contains("ВК") || docPath.Contains("АУПТ") || docPath.Contains("ЭОМ") || docPath.Contains("СС") || docPath.Contains("АВ")))
                {
                    if (docPath.Contains("ОБДН"))
                    {
                        gripBuilder = new GripBuilder_IOS(doc, "ОБДН", "SMNX_Этаж", 1, "SMNX_Секция", 0.328, 3);
                    }
                    else if (docPath.Contains("ИЗМЛ"))
                    {
                        gripBuilder = new GripBuilder_IOS(doc, "ИЗМЛ", "КП_О_Этаж", 1, "КП_О_Секция", 0.328, 10);
                    }
                }
                else
                {
                    TaskDialog taskDialog = new TaskDialog("ОШИБКА: Выполни инструкцию")
                    {
                        MainContent = "Ошибка номера отдела в БД. Обратись в BIM-отдел",
                        MainIcon = TaskDialogIcon.TaskDialogIconInformation
                    };
                    taskDialog.Show();
                    return Result.Failed;
                }

                if (gripBuilder == null)
                {
                    TaskDialog taskDialog = new TaskDialog("ОШИБКА: Выполни инструкцию")
                    {
                        MainContent = "Ошибка определения проекта. Обратись в BIM-отдел",
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
                        string errorIdColl = string.Join(
                            ",",
                            gripBuilder
                                .ErrorElements
                                .Where(e => e.ErrorMessage.Equals(error))
                                .Select(e => e.ErrorElement.Id.ToString()));

                        Print($"{error} - для след. элементов:\n {errorIdColl}", 
                            KPLN_Loader.Preferences.MessageType.Error);
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
            catch (Exception e)
            {
                Print($"Прервано с ошибкой: {e}", KPLN_Loader.Preferences.MessageType.Error);

                return Result.Failed;
            }
        }
    }
}