using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
                FilterRule rule = ParameterFilterRuleFactory
                    .CreateEqualsRule(userSelTypeParam.Id, userSelTypeName);
                resultFilter = new ElementParameterFilter(rule);
            }
            else
                throw new Exception($"Отправь разработчику: Не реализован поиск по параметру {bip} для эл-та: {userSelElem.Id}");

            return resultFilter;
        }

        internal static ElementParameterFilter SearchByParamName(Element userSelElem, string paramName)
        {
            ElementParameterFilter resultFilter;

            Parameter userSelTypeParam = userSelElem.LookupParameter(paramName);
            if (userSelTypeParam != null)
            {
                string userSelTypeName = userSelTypeParam.AsValueString();
                FilterRule rule = ParameterFilterRuleFactory
                    .CreateEqualsRule(userSelTypeParam.Id, userSelTypeName);
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
