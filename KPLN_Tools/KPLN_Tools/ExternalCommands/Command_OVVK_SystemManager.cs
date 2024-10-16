using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using KPLN_Tools.Common.OVVK_System;
using KPLN_Tools.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class Command_OVVK_SystemManager : IExternalCommand
    {
        /// <summary>
        /// Параметр толщины стенки труб
        /// </summary>
        private readonly Guid _thicknessParam = new Guid("381b467b-3518-42bb-b183-35169c9bdfb3");
#if Revit2020 || Debug2020
        private readonly DisplayUnitType _millimetersRevitType = DisplayUnitType.DUT_MILLIMETERS;
#endif
#if Revit2023 || Debug2023
        private readonly ForgeTypeId _millimetersRevitType = new ForgeTypeId("autodesk.unit.unit:millimeters-1.0.1");
#endif

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get application and documnet objects
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;



            return Result.Cancelled;
        }
    }
}
