using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_Parameters_Ribbon.Forms.Entities;
using System;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Parameters_Ribbon.Command
{
    public sealed class CommandUpdateParameterSums : IExecutableCommand
    {
        private readonly SumParametersM _sumParametersM;

        public CommandUpdateParameterSums(SumParametersM sumParametersM)
        {
            _sumParametersM = sumParametersM;
        }

        public Result Execute(UIApplication app)
        {
            try
            {
                UIDocument uidoc = app.ActiveUIDocument;
                if (uidoc == null)
                    return Result.Cancelled;

                _sumParametersM.SetUserSelection(uidoc.Document, SumParametersM.GetUserSelection(app));
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                PrintError(ex);
                return Result.Failed;
            }
        }
    }
}
