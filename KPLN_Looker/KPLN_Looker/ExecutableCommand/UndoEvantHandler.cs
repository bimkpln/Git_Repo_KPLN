using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Looker.ExecutableCommand
{
    internal class UndoEvantHandler : IExecutableCommand
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
