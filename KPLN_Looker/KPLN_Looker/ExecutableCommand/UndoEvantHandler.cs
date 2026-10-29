using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;

namespace KPLN_Looker.ExecutableCommand
{
    public class UndoEvantHandler : IExecutableCommand
    {
        private readonly Document _expectedDocument;

        public UndoEvantHandler() { }

        public UndoEvantHandler(Document expectedDocument)
        {
            _expectedDocument = expectedDocument;
        }

        public Result Execute(UIApplication app)
        {
            // Проверка DWG/PDF не должна отменять действие в другом документе.
            if (_expectedDocument != null
                && (!_expectedDocument.IsValidObject || !_expectedDocument.Equals(app.ActiveUIDocument?.Document)))
                throw new System.InvalidOperationException("Не удалось автоматически отменить DWG/PDF: активный документ изменился. Обратитесь в BIM-отдел.");

            try
            {
                app.PostCommand(RevitCommandId.LookupPostableCommandId(PostableCommand.Undo));
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException ioe)
            {
                if (_expectedDocument != null)
                    throw;

                if (!ioe.Message.Contains("Revit does not support more than one command are posted."))
                    throw;
            }

            return Result.Succeeded;
        }
    }
}
