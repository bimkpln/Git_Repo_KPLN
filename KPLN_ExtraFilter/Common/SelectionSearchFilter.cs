using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_ExtraFilter.Forms.Entities;
using System;
using System.Collections.Generic;
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

        internal static ElementParameterFilter SearchByParamName(Document doc, Element userSelElem, string paramName)
        {
            ElementParameterFilter resultFilter;

            Parameter userSelParam = userSelElem.LookupParameter(paramName);
            if (userSelParam == null && doc.GetElement(userSelElem.GetTypeId()) is Element typeElem)
                userSelParam = typeElem.LookupParameter(paramName);

            if (userSelParam != null)
            {
                FilterRule rule = null;
                switch (userSelParam.StorageType)
                {
                    case StorageType.ElementId:
                        rule = ParameterFilterRuleFactory
                            .CreateEqualsRule(userSelParam.Id, userSelParam.AsElementId());
                        break;
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
                throw new Exception($"Поиск по параметру {paramName} для эл-та: {userSelElem.Id} - НЕВОЗМОЖЕН. Скорее всего параметр отсутсвует у выбранного элемента");

            return resultFilter;
        }

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
