using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using KPLN_Parameters_Ribbon.Common.CheckParam.Builder;
using System;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Parameters_Ribbon.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandCheckParamData : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
#if Revit2023 || Revit2022 || Revit2020 || Revit2018
            TaskDialog.Show("упс....", "Ата-та... Еще работаем над этим!");
            return Result.Succeeded;
#endif


#if Debug2020 || Debug2023
            //Получение объектов приложения и документа
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            int userDepartment = Module.CurrentDBUser.SubDepartmentId;
            // Техническая подмена разделов для режима тестирования
            if (userDepartment == 6) { userDepartment = 4; }

            AbstrAuditBuilder auditBuilder = null;
            try
            {
                // Посадить на конфиг под каждый файл
                if (userDepartment == 1 || userDepartment == 4 && doc.Title.ToUpper().Contains("АР"))
                {
                    if (userDepartment == 1 || userDepartment == 4 && doc.Title.ToUpper().Contains("ОБДН"))
                    {
                        auditBuilder = new AuditBuilder_AR(doc, "ОБДН");
                    }
                }
                else if (userDepartment == 2 || userDepartment == 4 && doc.Title.ToUpper().Contains("КР"))
                {
                    if (userDepartment == 2 || userDepartment == 4 && doc.Title.ToUpper().Contains("ОБДН"))
                    {
                        auditBuilder = new AuditBuilder_AR(doc, "ОБДН");
                    }
                }
                else if (userDepartment == 3 || userDepartment == 4 && (doc.Title.ToUpper().Contains("ОВ") || doc.Title.ToUpper().Contains("ВК") || doc.Title.ToUpper().Contains("АУПТ") || doc.Title.ToUpper().Contains("ЭОМ") || doc.Title.ToUpper().Contains("СС") || doc.Title.ToUpper().Contains("АВ")))
                {
                    if (userDepartment == 3 || userDepartment == 4 && doc.Title.ToUpper().Contains("ОБДН"))
                    {
                        auditBuilder = new AuditBuilder_AR(doc, "ОБДН");
                    }
                }
                else
                {
                    throw new Exception("Ошибка номера отдела в БД. Обратись в BIM-отдел");
                }

                AuditDirector auditDirector = new AuditDirector(auditBuilder);
                auditDirector.BuildWriter();

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                Print($"Прервано с ошибкой: {e}", MessageType.Error);

                return Result.Failed;
            }
#endif
        }
    }
}
