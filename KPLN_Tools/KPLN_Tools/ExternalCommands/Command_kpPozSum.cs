using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Autodesk.Revit.UI.Selection;
using KPLN_Tools.Common.SS_System;
using KPLN_Tools.Forms;
using Newtonsoft.Json.Linq;
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

    internal class Command_kpPozSum : IExternalCommand
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

            using (Transaction t = new Transaction(_doc, "KPLN_Заполнение параметров"))
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

            return Result.Succeeded;
        }
    }
}
