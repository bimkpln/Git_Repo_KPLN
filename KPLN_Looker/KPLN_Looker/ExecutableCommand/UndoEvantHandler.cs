using Autodesk.Revit.UI;
using KPLN_Loader.Common;

namespace KPLN_Looker.ExecutableCommand
{
    public class UndoEvantHandler : IExecutableCommand
    {
        public Result Execute(UIApplication app)
        {
            try
            {
                app.PostCommand(RevitCommandId.LookupPostableCommandId(PostableCommand.Undo));
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException ioe)
            {
                if (!ioe.Message.Contains("Revit does not support more than one command are posted."))
                    throw ioe;
            }

            return Result.Succeeded;
        }
    }
}
