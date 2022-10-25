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
                if (userDepartment == 1 || userDepartment == 4 && doc.Title.ToUpper().Contains("АР"))
                {
                    if (userDepartment == 1 || userDepartment == 4 && doc.Title.ToUpper().Contains("ОБДН"))
                    {
                        gripBuilder = new GripBuilder_AR(doc, "ОБДН", "SMNX_Этаж", 1, "SMNX_Секция");
                    }
                }
                else if (userDepartment == 2 || userDepartment == 4 && doc.Title.ToUpper().Contains("КР"))
                {
                    if (userDepartment == 2 || userDepartment == 4 && doc.Title.ToUpper().Contains("ОБДН"))
                    {
                        gripBuilder = new GripBuilder_KR_OBDN(doc, "ОБДН", "SMNX_Этаж", 1, "SMNX_Секция");
                    }
                }
                else if (userDepartment == 3 || userDepartment == 4 && (doc.Title.ToUpper().Contains("ОВ") || doc.Title.ToUpper().Contains("ВК") || doc.Title.ToUpper().Contains("АУПТ") || doc.Title.ToUpper().Contains("ЭОМ") || doc.Title.ToUpper().Contains("СС") || doc.Title.ToUpper().Contains("АВ")))
                {
                    if (userDepartment == 3 || userDepartment == 4 && doc.Title.ToUpper().Contains("ОБДН"))
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
