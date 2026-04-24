using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_ExtraFilter.Forms.Entities;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace KPLN_ExtraFilter.Common
{
    /// <summary>
    /// Класс фильтрации Selection 
    /// </summary>
    internal class SelectorFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            // Фильтрация по классу (проемы)
            if (elem is Opening _)
                return false;

            // Фильтрация по категории (модельные эл-ты, кроме видов, сборок)
            if (elem.Category is Category elCat)
            {
                if (((elem.Category.CategoryType == CategoryType.Model)
                        || (elem.Category.CategoryType == CategoryType.Internal))
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                    && (elCat.Id.IntegerValue != (int)BuiltInCategory.OST_Viewers)
                    && (elCat.Id.IntegerValue != (int)BuiltInCategory.OST_IOSModelGroups)
                    && (elCat.Id.IntegerValue != (int)BuiltInCategory.OST_Assemblies))
#else
                    && (elCat.BuiltInCategory != BuiltInCategory.OST_Viewers)
                    && (elCat.BuiltInCategory != BuiltInCategory.OST_IOSModelGroups)
                    && (elCat.BuiltInCategory != BuiltInCategory.OST_Assemblies))
#endif
                    return true;
            }

            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            throw new NotImplementedException();
        }
    }


    /// <summary>
    /// Общий класс для генерации методов поиска
    /// </summary>
    internal static class SelectionSearchFilter
    {
        /// <summary>
        /// Метод для запуска рамки выбора
        /// </summary>
        /// <param name="uidoc"></param>
        /// <returns></returns>
        internal static IEnumerable<Element> UserSelectedFilters(UIDocument uidoc)
        {
            Document doc = uidoc.Document;

            try
            {
                IList<Reference> selectionRefers = uidoc.Selection.PickObjects(
                        ObjectType.Element,
                        new SelectorFilter(),
                        "Выберите нужные элементы (рамкой ИЛИ по одному) и нажмите \"Готово\"");

                List<Element> selElems = selectionRefers.Select(r => doc.GetElement(r.ElementId)).ToList();

                // Выделяю в модели
                uidoc.Selection.SetElementIds(selElems.Select(el => el.Id).ToList());
                
                return selElems;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return null;
            }
        }

        internal static ElementCategoryFilter SearchByCategory(Element userSelElem)
        {
            ElementCategoryFilter resultFilter;

#if Debug2020 || Revit2020 || Debug2023 || Revit2023
            BuiltInCategory bic = (BuiltInCategory)userSelElem.Category.Id.IntegerValue;
#else
            BuiltInCategory bic = userSelElem.Category.BuiltInCategory;
#endif
            if (bic != BuiltInCategory.INVALID)
                resultFilter = new ElementCategoryFilter(bic);
            else
                throw new Exception($"Отправь разработчику: Не реализован поиск по категории для эл-та: {userSelElem.Id}");

            return resultFilter;
        }

#if Debug2020 || Revit2020 || Debug2023 || Revit2023
        internal static ElementCategoryFilter SearchByCategoryEntity(CategoryEntity catEntity) => new ElementCategoryFilter((BuiltInCategory)catEntity.RevitCat.Id.IntegerValue);
#else
        internal static ElementCategoryFilter SearchByCategoryEntity(CategoryEntity catEntity) => new ElementCategoryFilter(catEntity.RevitCat.BuiltInCategory);
#endif

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

#elif !Debug2020 && !Revit2020
                FilterRule rule = ParameterFilterRuleFactory
                    .CreateEqualsRule(userSelTypeParam.Id, userSelTypeName);
#endif
                resultFilter = new ElementParameterFilter(rule);
            }
            else
                throw new Exception($"Отправь разработчику: Не реализован поиск по параметру {bip} для эл-та: {userSelElem.Id}");

            return resultFilter;
        }

        internal static ElementFilter SearchByParamName(Document doc, Element userSelElem, string paramName)
        {
            Parameter userSelParam = GetParameterFromElementOrType(doc, userSelElem, paramName) ?? throw new Exception(
                    $"Поиск по параметру {paramName} для эл-та: {userSelElem.Id} - НЕВОЗМОЖЕН. " +
                    $"Скорее всего параметр отсутствует у выбранного элемента");
            
            ElementId paramId = userSelParam.Id;

            if (!userSelParam.HasValue)
                return CreateFilterForEmptyParameter(paramId, userSelParam.StorageType);

            switch (userSelParam.StorageType)
            {
                case StorageType.ElementId:
                    return CreateFilter(ParameterFilterRuleFactory.CreateEqualsRule(paramId, userSelParam.AsElementId()));

                case StorageType.String:
                    string paramValue = userSelParam.AsString();

                    if (string.IsNullOrEmpty(paramValue))
                        return CreateStringEmptyOrNoValueFilter(paramId);

                    return CreateStringEqualsFilter(paramId, paramValue);

                case StorageType.Double:
                    return CreateFilter(ParameterFilterRuleFactory.CreateEqualsRule(paramId, userSelParam.AsDouble(), 0.01));

                case StorageType.Integer:
                    return CreateFilter(ParameterFilterRuleFactory.CreateEqualsRule(paramId, userSelParam.AsInteger()));

                default:
                    throw new Exception(
                        $"Поиск по параметру {paramName} для эл-та: {userSelElem.Id} - " +
                        $"не удалось создать фильтр. Unsupported StorageType: {userSelParam.StorageType}");
            }
        }

        private static Parameter GetParameterFromElementOrType(Document doc, Element element, string paramName)
        {
            Parameter param = element.LookupParameter(paramName);

            if (param == null && doc.GetElement(element.GetTypeId()) is Element typeElem)
                param = typeElem.LookupParameter(paramName);

            return param;
        }

        private static ElementFilter CreateFilterForEmptyParameter(ElementId paramId, StorageType storageType)
        {
            FilterRule noValueRule = ParameterFilterRuleFactory.CreateHasNoValueParameterRule(paramId);

            // Только текстовый параметр может быть либо без значения,
            // либо иметь пустую строку как значение
            if (storageType == StorageType.String)
                return CreateStringEmptyOrNoValueFilter(paramId);

            return CreateFilter(noValueRule);
        }

        private static ElementFilter CreateStringEmptyOrNoValueFilter(ElementId paramId)
        {
            ElementParameterFilter noValueFilter = CreateFilter(
                ParameterFilterRuleFactory.CreateHasNoValueParameterRule(paramId));

            ElementParameterFilter hasValueFilter = CreateFilter(
                ParameterFilterRuleFactory.CreateHasValueParameterRule(paramId));

            ElementParameterFilter emptyStringFilter = CreateStringEqualsFilter(paramId, string.Empty);

            // Просто emptyStringFilter использовать нельзя:
            // в некоторых случаях Revit начинает матчить слишком широко.
            ElementFilter fullEmptyValueFilter = new LogicalAndFilter(hasValueFilter, emptyStringFilter);

            return new LogicalOrFilter(noValueFilter, fullEmptyValueFilter);
        }

#if Debug2020 || Revit2020
        private static ElementParameterFilter CreateStringEqualsFilter(ElementId paramId, string value) => CreateFilter(ParameterFilterRuleFactory.CreateEqualsRule(paramId, value, false));
#else
        private static ElementParameterFilter CreateStringEqualsFilter(ElementId paramId, string value) => CreateFilter(ParameterFilterRuleFactory.CreateEqualsRule(paramId, value));
#endif

        private static ElementParameterFilter CreateFilter(FilterRule rule) => new ElementParameterFilter(rule);

        internal static ElementWorksetFilter SearchByElemWorkset(Element userSelElem)
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

        internal static ElementWorksetFilter SearchByWSEntity(WSEntity ws) => new ElementWorksetFilter(ws.RevitWSId);
    }
}
