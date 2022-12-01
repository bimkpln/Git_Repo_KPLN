using Autodesk.Revit.DB;
using KPLN_Classificator.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static KPLN_Classificator.ApplicationConfig;

namespace KPLN_Classificator
{
    public class ViewUtils
    {
        private static Color fullSuccessColor = new Color(255, 1, 1);
        private static Color notfullSuccessColor = new Color(1, 1, 255);
        public OverrideGraphicSettings fullSuccessSet;
        public OverrideGraphicSettings notFullSuccessSet;
        private Document doc;

        public ViewUtils(Document doc)
        {
            this.doc = doc;
            fullSuccessSet = getGraphicSettings(fullSuccessColor);
            notFullSuccessSet = getGraphicSettings(notfullSuccessColor);
        }

        private OverrideGraphicSettings getGraphicSettings(Color color)
        {
            FillPatternElement fillPattern = null;

            List<Element> fillPatterns = new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .WhereElementIsNotElementType()
                .ToElements()
                .ToList();

            foreach (var fp in fillPatterns)
            {
                if (fp.Name.ToString().Contains("Сплошная заливка"))
                {
                    fillPattern = (FillPatternElement)fp;
                    break;
                }
            }
            if (fillPattern == null)
            {
                output.PrintInfo("Заливка \"Сплошная заливка\" не найдена", Output.OutputMessageType.Error);
                return null;
            }

            OverrideGraphicSettings overrideGraphic = new OverrideGraphicSettings();
#if Revit2018
            overrideGraphic.SetProjectionFillPatternId(fillPattern.Id);
            overrideGraphic.SetProjectionFillColor(color);

#endif
#if Revit2020
            overrideGraphic.SetSurfaceForegroundPatternId(fillPattern.Id);
            overrideGraphic.SetSurfaceForegroundPatternColor(color);
#endif
            overrideGraphic.SetProjectionLineColor(color);

            return overrideGraphic;
        }

        public static OverrideGraphicSettings getStandartGraphicSettings(Document doc)
        {
            OverrideGraphicSettings overrideGraphic = new OverrideGraphicSettings();
#if Revit2018
            overrideGraphic.SetProjectionFillPatternId(ElementId.InvalidElementId);
            overrideGraphic.SetProjectionFillColor(Color.InvalidColorValue);
#endif
#if Revit2020
            overrideGraphic.SetSurfaceForegroundPatternId(ElementId.InvalidElementId);
            overrideGraphic.SetSurfaceForegroundPatternColor(Color.InvalidColorValue);
#endif
            overrideGraphic.SetProjectionLineColor(Color.InvalidColorValue);

            return overrideGraphic;
        }

        public static List<MyParameter> GetAllFilterableParameters(Document doc, List<Element> elements)
        {
            HashSet<ElementId> catsIds = GetElementsCategories(elements);

            Dictionary<ElementId, HashSet<ElementId>> paramsAndCats = new Dictionary<ElementId, HashSet<ElementId>>();

            foreach (ElementId catId in catsIds)
            {
                List<ElementId> curCatIds = new List<ElementId> { catId };
                List<ElementId> paramsIds = ParameterFilterUtilities.GetFilterableParametersInCommon(doc, curCatIds).ToList();

                foreach (ElementId paramId in paramsIds)
                {
                    if (paramsAndCats.ContainsKey(paramId))
                    {
                        paramsAndCats[paramId].Add(catId);
                    }
                    else
                    {
                        paramsAndCats.Add(paramId, new HashSet<ElementId> { catId });
                    }
                }
            }

            List<MyParameter> mparams = new List<MyParameter>();

            foreach (KeyValuePair<ElementId, HashSet<ElementId>> kvp in paramsAndCats)
            {
                ElementId paramId = kvp.Key;
                string paramName = GetParamName(doc, paramId);
                MyParameter mp = new MyParameter(paramId, paramName, kvp.Value.ToList());
                mparams.Add(mp);
            }

            mparams = mparams.OrderBy(i => i.Name).ToList();
            return mparams;
        }

        public static string GetCategoriesName(Document doc, List<ElementId> catIds0)
        {
            string result = "";

            List<ElementId> catIds = catIds0.Distinct().ToList();


            foreach (ElementId catId in catIds)
            {
                Category cat = Category.GetCategory(doc, catId);
                string catName = cat.Name;
                result += catName + ", ";
            }
            result = result.Substring(0, result.Length - 2);
            return result;
        }

        private static HashSet<ElementId> GetElementsCategories(List<Element> elements)
        {
            HashSet<ElementId> catsIds = new HashSet<ElementId>();
            foreach (Element elem in elements)
            {
                Category cat = elem.Category;
                if (cat == null) continue;
                if (cat.Id.IntegerValue == -2000500) continue;
                catsIds.Add(cat.Id);
            }
            return catsIds;
        }

        public static string GetParamName(Document doc, ElementId paramId)
        {
            string paramName = "error";
            if (paramId.IntegerValue < 0)
            {
                paramName = LabelUtils.GetLabelFor((BuiltInParameter)paramId.IntegerValue);
            }
            else
            {
                paramName = doc.GetElement(paramId).Name;
            }
            if (paramName != "error") return paramName;

            throw new Exception("Id не является идентификатором параметра: " + paramId.IntegerValue.ToString());
        }

    }
}
