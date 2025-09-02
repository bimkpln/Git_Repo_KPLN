using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_ExtensibleStorage;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;

namespace KPLN_ModelChecker_User.ExecutableCommand
{
    internal class CommandWPFEntity_SetTimeRunLog : IExecutableCommand
    {
        private readonly DBUser _currentDbUser;
        private readonly DateTime _closeTime;
        private readonly ExtensibleStorageBuilder _esBuilderRun;


        public CommandWPFEntity_SetTimeRunLog(ExtensibleStorageBuilder esBuilderRun, DateTime closeTime)
        {
            _esBuilderRun = esBuilderRun;
            _closeTime = closeTime;

            UserDbService userDbService = (UserDbService)new CreatorUserDbService().CreateService();
            _currentDbUser = userDbService.GetCurrentDBUser();
        }

        public Result Execute(UIApplication app)
        {
            // Игнорирую специалистов BIM-отдела
#if Revit2020 || Revit2023
            if (_currentDbUser.SubDepartmentId == 8) 
                return Result.Cancelled;
#endif

            using (Transaction t = new Transaction(app.ActiveUIDocument.Document, $"{ModuleData.ModuleName}_Время"))
            {
                t.Start();

                //Получение объектов приложения и документа
                Document doc = app.ActiveUIDocument.Document;
                ProjectInfo pi = doc.ProjectInformation;
                Element piElem = pi as Element;

                // Вписываю, если РН НЕ занят
                ICollection<ElementId> availableWSElemsId = WorksharingUtils.CheckoutElements(doc, new ElementId[] { piElem.Id });
                if (availableWSElemsId.Count > 0) 
                {
                    ExtensibleStorageBuilder esBuilder = new ExtensibleStorageBuilder(_esBuilderRun.Guid, _esBuilderRun.FieldName, _esBuilderRun.StorageName);
                    esBuilder.SetStorageData_TimeRunLog(piElem, app.Application.Username, _closeTime);
                }

                t.Commit();
            }

            return Result.Succeeded;
        }
    }
}
