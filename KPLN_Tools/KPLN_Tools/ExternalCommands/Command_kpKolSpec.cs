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
            Document _doc = commandData.Application.ActiveUIDocument.Document;

            FilteredElementCollector collector = new FilteredElementCollector(_doc);

            IList<Element> connectionElements = collector
                .OfCategory(BuiltInCategory.OST_StructConnections)
                .WhereElementIsNotElementType()
                .ToElements();

            List<Element> elemsMatched = new List<Element>();
            List<Element> elemsNotMatched = new List<Element>();

            foreach (Element elem in connectionElements)
            {
                ElementId typeId = elem.GetTypeId(); 
                ElementType elemType = _doc.GetElement(typeId) as ElementType; 

                if (elemType != null && (elemType.FamilyName == "076_КШ_Короб перфорированный_(ЭлУзл)" || elemType.FamilyName == "076_КШ_DIN рейка_(ЭлУзл)"))
                {
                    elemsMatched.Add(elem);
                }
                else
                {
                    elemsNotMatched.Add(elem);
                }
            }

            using (Transaction trans = new Transaction(_doc, "Обновление параметров"))
            {
                trans.Start();

                foreach (Element elem in elemsMatched)
                {
                    Parameter paramHeight = elem.LookupParameter("КП_Р_Высота");

                    if (paramHeight != null && paramHeight.HasValue)
                    {
                        string heightValue = paramHeight.AsString();

                        Parameter paramSpec = elem.LookupParameter("КП_И_КолСпецификация");

                        if (paramSpec != null && !paramSpec.IsReadOnly)
                        {
                            paramSpec.Set(heightValue);
                        }
                    }
                }

                foreach (Element elem in elemsNotMatched)
                {
                        Parameter paramSpec = elem.LookupParameter("КП_И_КолСпецификация");

                        if (paramSpec != null && !paramSpec.IsReadOnly)
                        {
                            paramSpec.Set(1);
                        }         
                }

                trans.Commit();
            }

            return Result.Succeeded;
        }
    }
}
