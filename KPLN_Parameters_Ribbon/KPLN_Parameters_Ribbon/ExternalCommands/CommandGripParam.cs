using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Parameters_Ribbon.Common.GripParam;
using KPLN_Parameters_Ribbon.Common.GripParam.Builder;
using KPLN_Parameters_Ribbon.Common.GripParam.Builder.OBDN;
using KPLN_Parameters_Ribbon.Forms;
using System;
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

            int userDepartment = KPLN_Loader.Preferences.User.Department.Id;
            // Техническая подмена оазделов для режима тестирования
            if (userDepartment == 6) { userDepartment = 4; }
            
            AbstrGripBuilder gripBuilder = null;
            try
            {
                // Посадить на конфиг под каждый файл
                string docTitle = doc.Title.ToUpper();
                if (userDepartment == 1 || userDepartment == 4 
                    && docTitle.Contains("АР"))
                {
                    if (docTitle.Contains("ОБДН"))
                    {
                        gripBuilder = new GripBuilder_AR(doc, "ОБДН", "SMNX_Этаж", 1, "SMNX_Секция");
                    }
                }
                else if (userDepartment == 2 || userDepartment == 4 
                    && docTitle.Contains("КР"))
                {
                    if (docTitle.Contains("ОБДН"))
                    {
                        gripBuilder = new GripBuilder_KR_OBDN(doc, "ОБДН", "SMNX_Этаж", 1, "SMNX_Секция");
                    }
                    else if (docTitle.Contains("ИЗМЛ"))
                    {
                        gripBuilder = new GripBuilder_KR(doc, "ИЗМЛ", "О_Этаж", 1, "КП_О_Секция");
                    }
                }
                else if (userDepartment == 3 || userDepartment == 4 
                    && (docTitle.Contains("ОВ") || docTitle.Contains("ВК") || docTitle.Contains("АУПТ") || docTitle.Contains("ЭОМ") || docTitle.Contains("СС") || docTitle.Contains("АВ")))
                {
                    if (docTitle.Contains("ОБДН"))
                    {
                        gripBuilder = new GripBuilder_IOS(doc, "ОБДН", "SMNX_Этаж", 1, "SMNX_Секция");
                    }
                }
                else
                {
                    throw new Exception("Ошибка номера отдела в БД. Обратись в BIM-отдел");
                }

                GripDirector gripDirector = new GripDirector(gripBuilder);
                gripDirector.BuildWriter();

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                Print($"Прервано с ошибкой: {e}", KPLN_Loader.Preferences.MessageType.Error);

                return Result.Failed;
            }
        }
    }
}
