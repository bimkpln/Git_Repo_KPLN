using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

using KPLN_Loader.Output;
using static KPLN_Finishing.Tools;
using static KPLN_Loader.Output.Output;
using System;
using System.Collections.Generic;
using System.Linq;
using KPLN_Finishing.Forms;
using static KPLN_Loader.Preferences;
using KPLN_Finishing.CommandTools;

namespace KPLN_Finishing.ExternalCommands
{
    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]
    class LoadParameters : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            try
            {
                LoadRoomParameters(doc);
                LoadElementParameters(doc);
            }
            catch (Exception e)
            {
                Output.PrintError(e);
                return Result.Failed;
            }
            return Result.Succeeded;
        }
    }
}
