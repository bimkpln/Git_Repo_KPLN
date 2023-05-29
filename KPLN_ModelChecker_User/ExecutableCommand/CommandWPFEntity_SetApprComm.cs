using Autodesk.Revit.UI;
using KPLN_Library_ExtensibleStorage;
using KPLN_Loader.Common;
using KPLN_ModelChecker_User.WPFItems;

namespace KPLN_ModelChecker_User.ExecutableCommand
{
    internal class CommandWPFEntity_SetApprComm : IExecutableCommand
    {
        private WPFEntity _wpfEntity;
        private ExtensibleStorageBuilder _esText;
        private string _description;

        public CommandWPFEntity_SetApprComm(WPFEntity wpfEntity, ExtensibleStorageBuilder esText, string description)
        {
            _wpfEntity = wpfEntity;
            _esText = esText;
            _description = description;
        }

        public Result Execute(UIApplication app)
        {
            //Получение объектов приложения и документа
            ExtensibleStorageBuilder esBuilder = new ExtensibleStorageBuilder(_esText.Guid, _esText.FieldName, _esText.StorageName);
            esBuilder.SetStorageData_TextLog(_wpfEntity.Element, app.Application.Username, _description);

            //Обновление данных на wpf-элементе
            _wpfEntity.UpdateMainFieldByStatus(Common.Collections.Status.Approve);
            _wpfEntity.ApproveComment = _description;

            return Result.Succeeded;
        }
    }
}
