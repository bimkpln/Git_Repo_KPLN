using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Autodesk.Revit.UI.Selection;
using KPLN_Tools.Common.SS_System;
using KPLN_Tools.Forms;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class Command_kpKolSpec : IExternalCommand
    {     
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication _uiapp = commandData.Application;
            UIDocument _uidoc = _uiapp.ActiveUIDocument;
            Document _doc = _uidoc.Document;

            var familyInstances = new FilteredElementCollector(_doc, _doc.ActiveView.Id)
            .OfClass(typeof(FamilyInstance))
            .Cast<FamilyInstance>()
            .Where(fi =>
                fi.Symbol.FamilyName == "076_КШ_Короб перфорированный_(ЭлУзл)" ||
                fi.Symbol.FamilyName == "076_КШ_DIN рейка_(ЭлУзл)");

            double totalLength = 0;

            foreach (var familyInstance in familyInstances)
            {
                Parameter groupingParam = familyInstance.LookupParameter("КП_О_Группирование");

                if (groupingParam != null && groupingParam.AsInteger() == 1) 
                {
                    Parameter lengthParam = familyInstance.LookupParameter("Длина");
                    if (lengthParam != null)
                    {
                        totalLength += lengthParam.AsDouble();
                    }
                }
            }

            using (Transaction trans = new Transaction(_doc, "Запись длины в КП_И_КолСпецификация"))
            {
                trans.Start();
                double lengthInMeters = UnitUtils.ConvertFromInternalUnits(totalLength, UnitTypeId.Meters);

                foreach (var familyInstance in familyInstances)
                {
                    Parameter specParam = familyInstance.LookupParameter("КП_И_КолСпецификация");
                    if (specParam != null)
                    {
                        specParam.Set(lengthInMeters);
                    }
                }

                trans.Commit();
            }

            return Result.Succeeded;
        }
    }
}
