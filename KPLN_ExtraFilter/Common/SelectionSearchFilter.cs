using Autodesk.Revit.DB;
using System;
using System.Linq;

namespace KPLN_ExtraFilter.Common
{
    /// <summary>
    /// Общий класс для генерации методов поиска
    /// </summary>
    internal static class SelectionSearchFilter
    {
        internal static ElementCategoryFilter SearchByCategory(Element userSelElem)
        {
            ElementCategoryFilter resultFilter;

            BuiltInCategory bic = (BuiltInCategory)userSelElem.Category.Id.IntegerValue;
            if (bic != BuiltInCategory.INVALID)
                resultFilter = new ElementCategoryFilter((BuiltInCategory)userSelElem.Category.Id.IntegerValue);
            else
                throw new Exception($"Отправь разработчику: Не реализован поиск по категории для эл-та: {userSelElem.Id}");

            return resultFilter;
        }

        internal static ElementParameterFilter SearchByElemBuiltInParam(Element userSelElem, BuiltInParameter bip)
        {
            ElementParameterFilter resultFilter;

            Parameter userSelTypeParam = userSelElem.get_Parameter(bip);
            if (userSelTypeParam != null)
            {
                string userSelTypeName = userSelTypeParam.AsValueString();
#if Debug2020 || Revit2020
                FilterRule rule = ParameterFilterRuleFactory
                    .CreateEqualsRule(userSelTypeParam.Id, userSelTypeName, false);

#elif Debug2023 || Revit2023
                FilterRule rule = ParameterFilterRuleFactory
                    .CreateEqualsRule(userSelTypeParam.Id, userSelTypeName);
#endif
                resultFilter = new ElementParameterFilter(rule);
            }
            else
                throw new Exception($"Отправь разработчику: Не реализован поиск по параметру {bip} для эл-та: {userSelElem.Id}");

            return resultFilter;
        }

        internal static ElementParameterFilter SearchByParamName(Element userSelElem, string paramName)
        {
            ElementParameterFilter resultFilter;

            Parameter userSelParam = userSelElem.LookupParameter(paramName);
            if (userSelParam != null)
            {
                FilterRule rule = null;
                switch (userSelParam.StorageType)
                {
                    case StorageType.String:
#if Debug2020 || Revit2020
                        rule = ParameterFilterRuleFactory
                            .CreateEqualsRule(userSelParam.Id, userSelParam.AsString(), false);
#elif Debug2023 || Revit2023
                        rule = ParameterFilterRuleFactory
                            .CreateEqualsRule(userSelParam.Id, userSelParam.AsString());
#endif
                        break;
                    case StorageType.Double:
                        rule = ParameterFilterRuleFactory
                            .CreateEqualsRule(userSelParam.Id, userSelParam.AsDouble(), 0.01);
                        break;
                    case StorageType.Integer:
                        rule = ParameterFilterRuleFactory
                            .CreateEqualsRule(userSelParam.Id, userSelParam.AsInteger());
                        break;
                }

                resultFilter = new ElementParameterFilter(rule);
            }
            else
                throw new Exception($"Отправь разработчику: Не реализован поиск по параметру {paramName} для эл-та: {userSelElem.Id}");

            return resultFilter;
        }

        internal static ElementWorksetFilter SearchByWorkset(Element userSelElem)
        {
            ElementWorksetFilter resultFilter;

            Parameter elemWSParam = userSelElem.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
            if (elemWSParam != null)
            {
                string elemWSName = elemWSParam.AsValueString();
                Workset workset = new FilteredWorksetCollector(userSelElem.Document)
                    .OfKind(WorksetKind.UserWorkset)
                    .FirstOrDefault(ws => ws.Name.Equals(elemWSName));

                resultFilter = new ElementWorksetFilter(workset.Id);
            }
            else
                throw new Exception(
                    $"Отправь разработчику: Не удалось реализовать поиск по рабочему набору для эл-та: {userSelElem.Id}");

            return resultFilter;
        }
    }
}
