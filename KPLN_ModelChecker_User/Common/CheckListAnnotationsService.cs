using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib;
using KPLN_ModelChecker_Lib.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using MediaColor = System.Windows.Media.Color;

namespace KPLN_ModelChecker_User.Common
{
    public static class CheckListAnnotationsService
    {
        private sealed class HighlightedFilter
        {
            public ElementId FilterId { get; set; }
            
            public List<ElementId> ViewIds { get; set; } = new List<ElementId>();
        }

        private const string FilterPrefix = "KPLN_AnnotationsHighlight";
        private static readonly Dictionary<string, List<HighlightedFilter>> _highlightedFilters = new Dictionary<string, List<HighlightedFilter>>();

        public static string ApplyHighlightOrSelection(
            CheckListAnnotationsSettingsM selConfig,
            UIApplication uiapp,
            AbstrCheck abstrCheck)
        {
            string resultStatus = string.Empty;
            
            
            if (selConfig == null 
                || uiapp?.ActiveUIDocument?.Document == null 
                || abstrCheck.CheckerEntitiesColl == null)
                return string.Empty;

            Document doc = uiapp.ActiveUIDocument.Document;

            // Запуск сценария выделения элементов в модели, если выбран соответствующий режим отображения
            if (selConfig.DisplayMode == CheckListAnnotationsDisplayMode.SelectElements)
            {
                uiapp.ActiveUIDocument.Selection.SetElementIds(abstrCheck.CheckerEntitiesColl.SelectMany(ent => ent.ElementIdCollection).ToList());
                resultStatus = $"Выделено {abstrCheck.CheckerEntitiesColl.Sum(ce => ce.ElementIdCollection.Count())} элементов.";
            }


            // Очистка от старых переопределений графики, если выбрана соответствующая опция
            if (selConfig.ClearPreviousHighlight)
                ClearCurrentHighlight(uiapp, false);


            // Обрабатываю виды, на которых находятся элементы, и создаю фильтры для подсветки
            List<ElementId> elementIds = GetElementIds(abstrCheck.CheckerEntitiesColl);
            Dictionary<ElementId, List<ElementId>> targetViewIdsDict = GetTargetViews(doc, elementIds);
            string key = GetDocumentKey(doc);

            
            // Пополняю словарь для обнуления
            if (!_highlightedFilters.ContainsKey(key))
                _highlightedFilters[key] = new List<HighlightedFilter>();


            // Основная часть
            if (targetViewIdsDict.Keys.Any())
            {
                List<ElementId> elemsToSelect = new List<ElementId>();
                List<ElementId> elemsToHighlight = new List<ElementId>();
                using (Transaction transaction = new Transaction(doc, $"{ModuleData.ModuleName}_Подсветка аннотаций"))
                {
                    transaction.Start();

                    SelectionFilterElement filterElement = SelectionFilterElement.Create(doc, CreateFilterName());
                    filterElement.SetElementIds(elementIds);

                    OverrideGraphicSettings overrideSettings = CreateOverrideSettings(selConfig.HighlightColor);
                    HighlightedFilter highlightedFilter = new HighlightedFilter { FilterId = filterElement.Id };

                    
                    foreach (var kvp in targetViewIdsDict)
                    {
                        var viewId = kvp.Key;
                        if (!(doc.GetElement(viewId) is View view))
                            continue;

                        // Обработка листов
                        if (view is ViewSheet _)
                        {
                            elemsToSelect.AddRange(kvp.Value);
                        }
                        // Обработка видов
                        else
                        {
                            if (!EnableTemporaryViewPropertiesMode(view))
                                continue;

                            try
                            {
                                if (!view.GetFilters().Contains(filterElement.Id))
                                    view.AddFilter(filterElement.Id);

                                view.SetFilterOverrides(filterElement.Id, overrideSettings);
                                highlightedFilter.ViewIds.Add(view.Id);
                            }
                            catch
                            {
                            }

                            elemsToHighlight.AddRange(kvp.Value);
                        }

                    }

                    if (highlightedFilter.ViewIds.Count > 0)
                        _highlightedFilters[key].Add(highlightedFilter);
                    else
                        doc.Delete(filterElement.Id);

                    transaction.Commit();
                }

                // Выделяю элементы на листах (их фильтрами не переопределить)
                if (elemsToSelect.Any())
                    uiapp.ActiveUIDocument.Selection.SetElementIds(elemsToSelect);

                // Заполняю отчет для пользователя
                if (elemsToSelect.Any() && elemsToHighlight.Any())
                    resultStatus = $"Выделено на листе {elemsToSelect.Count()} шт., окрашено на виде/-ах: {elemsToHighlight.Count()} шт.";
                else if(elemsToSelect.Any())
                    resultStatus = $"Выделено на листе {elemsToSelect.Count()} шт.";
                else if (elemsToHighlight.Any())
                    resultStatus = $"Окрашено на виде/-ах: {elemsToHighlight.Count()} шт.";


                return resultStatus;
            }

            return string.Empty;
        }

        /// <summary>
        /// Снимает текущую подсветку аннотаций в модели, удаляя фильтры и отключая временный режим отображения для видов.
        /// </summary>
        /// <param name="uiapp"></param>
        /// <param name="showResult"></param>
        public static void ClearCurrentHighlight(UIApplication uiapp, bool showResult = true)
        {
            if (uiapp?.ActiveUIDocument?.Document == null)
                return;

            Document doc = uiapp.ActiveUIDocument.Document;
            string key = GetDocumentKey(doc);
            List<HighlightedFilter> highlightedFilters = GetHighlightedFilters(doc, key);

            if (highlightedFilters.Count == 0)
            {
                if (showResult)
                    TaskDialog.Show("KPLN", "Текущая подсветка проверки аннотаций не найдена.", TaskDialogCommonButtons.Ok);

                return;
            }

            List<ElementId> viewIds = highlightedFilters
                .SelectMany(filter => filter.ViewIds)
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                .GroupBy(id => id.IntegerValue)
#else
                .GroupBy(id => id.Value)
#endif
                .Select(group => group.First())
                .ToList();

            int clearedViewCount = 0;
            int deletedFilterCount = 0;

            using (Transaction transaction = new Transaction(doc, $"{ModuleData.ModuleName}_Снять подсветку аннотаций"))
            {
                transaction.Start();

                foreach (ElementId viewId in viewIds)
                {
                    if (!(doc.GetElement(viewId) is View view) || !view.IsValidObject)
                        continue;

                    try
                    {
                        if (view.IsTemporaryViewPropertiesModeEnabled())
                        {
                            view.DisableTemporaryViewMode(TemporaryViewMode.TemporaryViewProperties);
                            clearedViewCount++;
                        }
                    }
                    catch
                    {
                    }
                }

                foreach (HighlightedFilter highlightedFilter in highlightedFilters)
                {
                    try
                    {
                        Element filterElement = doc.GetElement(highlightedFilter.FilterId);
                        if (filterElement != null && filterElement.IsValidObject)
                        {
                            doc.Delete(highlightedFilter.FilterId);
                            deletedFilterCount++;
                        }
                    }
                    catch
                    {
                    }
                }

                transaction.Commit();
            }

            _highlightedFilters[key] = new List<HighlightedFilter>();

            if (showResult)
                TaskDialog.Show("KPLN", $"Подсветка снята. Отмена переопределения видов для {clearedViewCount} шт.; удалено фильтров {deletedFilterCount} шт.", TaskDialogCommonButtons.Ok);
        }

        /// <summary>
        /// Включает временный режим отображения свойств для указанного вида, если он еще не включен.
        /// </summary>
        /// <param name="view"></param>
        /// <returns></returns>
        private static bool EnableTemporaryViewPropertiesMode(View view)
        {
            if (view == null || !view.IsValidObject)
                return false;

            try
            {
                if (view.IsTemporaryViewPropertiesModeEnabled())
                    return true;

                return view.EnableTemporaryViewPropertiesMode(view.Id);
            }
            catch
            {
                return false;
            }
        }

        private static List<HighlightedFilter> GetHighlightedFilters(Document doc, string key)
        {
            Dictionary<ElementId, HighlightedFilter> result = new Dictionary<ElementId, HighlightedFilter>();

            if (_highlightedFilters.ContainsKey(key))
            {
                foreach (HighlightedFilter highlightedFilter in _highlightedFilters[key])
                    AddHighlightedFilter(result, highlightedFilter.FilterId, highlightedFilter.ViewIds);
            }

            List<View> temporaryViews = GetTemporaryViewPropertiesViews(doc);
            IEnumerable<SelectionFilterElement> documentFilters = new FilteredElementCollector(doc)
                .OfClass(typeof(SelectionFilterElement))
                .Cast<SelectionFilterElement>()
                .Where(filter => filter.Name.StartsWith(FilterPrefix));

            foreach (SelectionFilterElement filterElement in documentFilters)
            {
                List<ElementId> viewIds = temporaryViews
                    .Where(view => ViewContainsFilter(view, filterElement.Id))
                    .Select(view => view.Id)
                    .ToList();

                AddHighlightedFilter(result, filterElement.Id, viewIds);
            }

            return result.Values.ToList();
        }

        private static void AddHighlightedFilter(
            Dictionary<ElementId, HighlightedFilter> highlightedFilters,
            ElementId filterId,
            IEnumerable<ElementId> viewIds)
        {
            if (!highlightedFilters.ContainsKey(filterId))
                highlightedFilters.Add(filterId, new HighlightedFilter { FilterId = filterId });

            foreach (ElementId viewId in viewIds)
            {
                if (!highlightedFilters[filterId].ViewIds.Contains(viewId))
                    highlightedFilters[filterId].ViewIds.Add(viewId);
            }
        }

        private static List<View> GetTemporaryViewPropertiesViews(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(view => view != null && view.IsValidObject && IsTemporaryViewPropertiesModeEnabled(view))
                .ToList();
        }

        private static bool IsTemporaryViewPropertiesModeEnabled(View view)
        {
            try
            {
                return view.IsTemporaryViewPropertiesModeEnabled();
            }
            catch
            {
                return false;
            }
        }

        private static bool ViewContainsFilter(View view, ElementId filterId)
        {
            try
            {
                return view.GetFilters().Contains(filterId);
            }
            catch
            {
                return false;
            }
        }

        private static List<ElementId> GetElementIds(IEnumerable<CheckerEntity> checkerEntities)
        {
            return checkerEntities
                .Where(entity => entity?.ElementIdCollection != null)
                .SelectMany(entity => entity.ElementIdCollection)
                .Where(id => id != null && !id.Equals(ElementId.InvalidElementId))
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                .GroupBy(id => id.IntegerValue)
#else
                .GroupBy(id => id.Value)
#endif
                .Select(group => group.First())
                .ToList();
        }

        /// <summary>
        /// Возвращает список видов с привязкой к списку элементов, на которых находятся элементы с указанными идентификаторами.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="elementIds"></param>
        /// <returns></returns>
        private static Dictionary<ElementId, List<ElementId>> GetTargetViews(Document doc, IEnumerable<ElementId> elementIds)
        {
            Dictionary <ElementId, List< ElementId >> result = new Dictionary<ElementId, List<ElementId>>();

            foreach (ElementId elementId in elementIds)
            {
                var elem = doc.GetElement(elementId);

                if (elem == null || !elem.IsValidObject)
                    continue;

                var viewId = elem.OwnerViewId.Equals(ElementId.InvalidElementId)
                    ? doc.ActiveView.Id
                    : elem.OwnerViewId;

                if (!(doc.GetElement(viewId) is View view) || !view.IsValidObject)
                    continue;

                if (!result.ContainsKey(viewId))
                    result[viewId] = new List<ElementId>();

                result[viewId].Add(elementId);
            }

            return result;
        }

        /// <summary>
        /// Создает объект OverrideGraphicSettings с указанным цветом для переопределения графики элементов.
        /// </summary>
        /// <param name="color"></param>
        /// <returns></returns>
        private static OverrideGraphicSettings CreateOverrideSettings(MediaColor color)
        {
            OverrideGraphicSettings settings = new OverrideGraphicSettings();
            Autodesk.Revit.DB.Color revitColor = new Autodesk.Revit.DB.Color(color.R, color.G, color.B);
            settings.SetProjectionLineColor(revitColor);
            settings.SetCutLineColor(revitColor);
            
            return settings;
        }

        private static string CreateFilterName() => $"{FilterPrefix}_{Guid.NewGuid():N}";

        /// <summary>
        /// Возвращает уникальный ключ для документа, используя его путь или заголовок.
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        private static string GetDocumentKey(Document doc)
        {
            if (!string.IsNullOrWhiteSpace(doc.PathName))
                return doc.PathName;

            return doc.Title;
        }
    }
}