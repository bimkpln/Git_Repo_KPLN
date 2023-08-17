using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_ExtensibleStorage;
using KPLN_Loader.Common;
using System;

namespace KPLN_ModelChecker_User.ExecutableCommand
{
    internal class CommandWPFEntity_SetTimeRunLog : IExecutableCommand
    {
        private readonly DateTime _closeTime;
        private readonly ExtensibleStorageBuilder _esBuilderRun;

        public CommandWPFEntity_SetTimeRunLog(ExtensibleStorageBuilder esBuilderRun, DateTime closeTime)
        {
            _esBuilderRun = esBuilderRun;
            _closeTime = closeTime;
        }

        public Result Execute(UIApplication app)
        {
            using (Transaction t = new Transaction(app.ActiveUIDocument.Document, $"{ModuleData.ModuleName}_Время"))
            {
                t.Start();

                // Игнорирую специалистов BIM-отдела
                int _userDepartment = KPLN_Loader.Application.CurrentRevitUser.SubDepartmentId;
                if (_userDepartment == 8) return Result.Cancelled;

                //Получение объектов приложения и документа
                Document doc = app.ActiveUIDocument.Document;
                ProjectInfo pi = doc.ProjectInformation;
                Element piElem = pi as Element;
                ExtensibleStorageBuilder esBuilder = new ExtensibleStorageBuilder(_esBuilderRun.Guid, _esBuilderRun.FieldName, _esBuilderRun.StorageName);
                esBuilder.SetStorageData_TimeRunLog(piElem, app.Application.Username, _closeTime);

                t.Commit();
            }

            return Result.Succeeded;
        }
    }
}
