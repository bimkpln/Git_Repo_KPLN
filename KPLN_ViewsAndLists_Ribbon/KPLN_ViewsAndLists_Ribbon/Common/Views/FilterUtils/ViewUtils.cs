﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using KPLN_ViewsAndLists_Ribbon.Views.Colorize;

namespace KPLN_ViewsAndLists_Ribbon.Views.FilterUtils
{
    public static class ViewUtils
    {
        public static void ApplyViewFilter(Document doc, View view, ParameterFilterElement filter, ElementId solidFillPatternId, int colorNumber, bool colorLines, bool colorFill)
        {
            view.AddFilter(filter.Id);
            OverrideGraphicSettings ogs = new OverrideGraphicSettings();

            byte red = Convert.ToByte(ColorsCollection.colors[colorNumber].Substring(1, 2), 16);
            byte green = Convert.ToByte(ColorsCollection.colors[colorNumber].Substring(3, 2), 16);
            byte blue = Convert.ToByte(ColorsCollection.colors[colorNumber].Substring(5, 2), 16);

            Color clr = new Color(red, green, blue);

            if (colorLines)
            {
                ogs.SetProjectionLineColor(clr);
                ogs.SetCutLineColor(clr);
            }

            if (colorFill)
            {

#if R2017 || R2018
            ogs.SetProjectionFillColor(clr);
            ogs.SetProjectionFillPatternId(solidFillPatternId);
            ogs.SetCutFillColor(clr);
            ogs.SetCutFillPatternId(solidFillPatternId);
#else
                ogs.SetSurfaceForegroundPatternColor(clr);
                ogs.SetSurfaceForegroundPatternId(solidFillPatternId);
                ogs.SetCutForegroundPatternColor(clr);
                ogs.SetCutForegroundPatternId(solidFillPatternId);
#endif
            }
            view.SetFilterOverrides(filter.Id, ogs);
        }


        public static bool CheckIsChangeFiltersAvailable(Document doc, View view)
        {
            ElementId templateId = view.ViewTemplateId;
            if (templateId == null) return true;
            if (templateId == ElementId.InvalidElementId) return true;


            View template = doc.GetElement(templateId) as View;

            List<int> nonControlledParmsIds = template.GetNonControlledTemplateParameterIds()
                .Select(p => p.IntegerValue)
                .ToList();

            List<int> allTemplateParams = template.GetTemplateParameterIds()
                .Select(p => p.IntegerValue)
                .ToList();

            int visFiltersParamId = (int)BuiltInParameter.VIS_GRAPHICS_FILTERS;

            foreach (int id in allTemplateParams)
            {
                if (id != visFiltersParamId) continue;

                if (nonControlledParmsIds.Contains(id))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            throw new Exception("Не удалось определить возможность применения фильтров для вида " + view.Name);
        }


        /// <summary>
        /// Полный список доступных для фильтрации параметров у каждой категории
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="elements"></param>
        /// <returns></returns>
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
