using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_ExtensibleStorage;
using KPLN_Loader.Common;
using System;

namespace KPLN_ModelChecker_User.ExecutableCommand
{
    internal class CommandWPFEntity_SetTimeRunLog : IExecutableCommand
    {
        private DateTime _closeTime;
        private ExtensibleStorageBuilder _esBuilderRun;

        public CommandWPFEntity_SetTimeRunLog(ExtensibleStorageBuilder esBuilderRun, DateTime closeTime)
        {
            _esBuilderRun = esBuilderRun;
            _closeTime = closeTime;
        }

        public Result Execute(UIApplication app)
        {
            // Игнорирую специалистов BIM-отдела
            int _userDepartment = KPLN_Loader.Preferences.User.Department.Id;
            if (_userDepartment == 4)
            {
                return Result.Cancelled;
            }

            //Получение объектов приложения и документа
            Document doc = app.ActiveUIDocument.Document;
            ProjectInfo pi = doc.ProjectInformation;
            Element piElem = pi as Element;
            ExtensibleStorageBuilder esBuilder = new ExtensibleStorageBuilder(_esBuilderRun.Guid, _esBuilderRun.FieldName, _esBuilderRun.StorageName);
            esBuilder.SetStorageData_TimeRunLog(piElem, app.Application.Username, _closeTime);

            return Result.Succeeded;
        }
    }
}
