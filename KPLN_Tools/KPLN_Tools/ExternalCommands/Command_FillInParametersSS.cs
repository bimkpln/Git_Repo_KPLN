using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]

    internal class Command_FillInParametersSS : IExternalCommand
    {
        public Parameter GetParameterByName(Document doc, Element element, string parameterName)
        {
            Parameter parameter = element.LookupParameter(parameterName);

            if (parameter == null)
            {
                parameter = doc
                    .GetElement(element.GetTypeId())
                    .LookupParameter(parameterName);
            }

            return parameter;
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication _uiapp = commandData.Application;
            UIDocument _uidoc = _uiapp.ActiveUIDocument;
            Document _doc = _uidoc.Document;

            // Заполнение параметров: КП_Позиция_Сумма
            IList<Element> allElementsInProject = new FilteredElementCollector(_doc)
            .OfCategory(BuiltInCategory.OST_DetailComponents)
            .WhereElementIsNotElementType()
            .ToElements();

            Dictionary<string, (string dValue, List<Element> dPair)> newDict = new Dictionary<string, (string, List<Element>)>();

            foreach (Element element in allElementsInProject)
            {
                Parameter keyParam = GetParameterByName(_doc, element, "КП_О_Наименование");
                Parameter key2Param = GetParameterByName(_doc, element, "КП_О_Группирование");
                Parameter valueParam = GetParameterByName(_doc, element, "КП_О_Позиция");

                if (keyParam == null || key2Param == null || valueParam == null) continue;

                string key = keyParam.AsString();
                string key2 = key2Param.AsString();
                string value = valueParam.AsString();

                if (string.IsNullOrEmpty(key) || string.IsNullOrEmpty(key2) || string.IsNullOrEmpty(value)) continue;

                string keyVal = key + key2;

                if (!newDict.ContainsKey(keyVal))
                {
                    newDict[keyVal] = (value + "\\", new List<Element> { element });
                }
                else
                {
                    var existing = newDict[keyVal];
                    existing.dValue += value + "\\";
                    existing.dPair.Add(element);
                    newDict[keyVal] = existing;
                }
            }
   
            using (Transaction t = new Transaction(_doc, "KPLN_Заполнение параметров: КП_Позиция_Сумма"))
            {
                t.Start();

                foreach (var entry in newDict)
                {
                    string value = entry.Value.dValue;
                    var symbols = value.Split('\\').Distinct().OrderBy(sym => sym).ToList();

                    string newValue = string.Join("; ", symbols.Where(sym => !string.IsNullOrEmpty(sym)));
                    if (string.IsNullOrEmpty(newValue))
                        continue;

                    foreach (Element parElement in entry.Value.dPair)
                    {
                        Parameter paramToSet = GetParameterByName(_doc, parElement, "КП_Позиция_Сумма");
                        if (paramToSet != null)
                        {
                            paramToSet.Set(newValue);
                        }
                    }
                }

                t.Commit();
            }

            // Заполнение параметров: КП_И_КолСпецификация
            List<Element> connectionElements = new List<Element>();
            FilteredElementCollector collector = new FilteredElementCollector(_doc);

            foreach (Element e in collector.WhereElementIsNotElementType())
            {
                if (e.Category != null && e.Category.Name == "Элементы узлов")
                {
                    connectionElements.Add(e);
                }
            }

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

            using (Transaction t2 = new Transaction(_doc, "KPLN_Заполнение параметров: КП_И_КолСпецификация"))
            {
                t2.Start();

                foreach (Element elem in elemsMatched)
                {
                    Parameter paramHeight = elem.LookupParameter("КП_Р_Высота");

                    if (paramHeight != null && paramHeight.HasValue)
                    {
                        double heightValue = paramHeight.AsDouble() * 304.8;
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

                t2.Commit();
            }

            MessageBox.Show("Все параметры были заполнены", "Результат", MessageBoxButtons.OK, MessageBoxIcon.Information);
      
            return Result.Succeeded;
        }
    }
}
