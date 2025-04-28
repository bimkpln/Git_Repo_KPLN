using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class Command_OVVK_SystemManager : IExternalCommand
    {
        internal const string PluginName = "Менеджер систем";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get application and documnet objects
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            
            OVVK_SystemManagerForm form = new OVVK_SystemManagerForm(doc);
            form.Show();

            return Result.Succeeded;
        }
    }
}
