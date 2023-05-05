using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_ExtensibleStorage;
using KPLN_Loader.Common;

namespace KPLN_ModelChecker_User.ExecutableCommand
{
    internal class CommandSetExtStr_TextLog : IExecutableCommand
    {
        private Element _element;
        private ExtensibleStorageBuilder _esText;
        private string _description;

        public CommandSetExtStr_TextLog(Element element, ExtensibleStorageBuilder esText, string description)
        {
            _element = element;
            _esText = esText;
            _description = description;
        }

        public Result Execute(UIApplication app)
        {
            //Получение объектов приложения и документа
            ExtensibleStorageBuilder esBuilder = new ExtensibleStorageBuilder(_esText.Guid, _esText.FieldName, _esText.StorageName);
            esBuilder.SetStorageData_TextLog(_element, app.Application.Username, _description);

            return Result.Succeeded;
        }
    }
}
