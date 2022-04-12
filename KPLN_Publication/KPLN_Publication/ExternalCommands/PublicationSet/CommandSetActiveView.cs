using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static KPLN_Loader.Output.Output;

namespace KPLN_Publication.ExternalCommands.PublicationSet
{
    public class CommandSetActiveView : IExecutableCommand
    {
        View View { get; set; }
        public CommandSetActiveView(View view)
        {
            View = view;
        }
        public Result Execute(UIApplication app)
        {
            //Print("Execute CommandSetActiveView", KPLN_Loader.Preferences.MessageType.System_Regular);
            try
            {
                app.ActiveUIDocument.RequestViewChange(View);
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                PrintError(e);
                return Result.Failed;
            }
        }
    }
}
