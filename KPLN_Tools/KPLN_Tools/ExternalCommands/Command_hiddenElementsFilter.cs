using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Common;
using KPLN_Tools.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Threading;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class Command_hiddenElementsFilter : IExternalCommand
    {
        internal const string PluginName = "Управление скрытыми элементами";
        private const string FilterParameterName = "KPLN_Фильтрация";
        private const string FilterNamePrefix = "KPLN_Фильтрация_";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            ScanProgressWindow progressWindow = null;

            try
            {
                progressWindow = new ScanProgressWindow();
                progressWindow.Show();
                progressWindow.UpdateProgress(0, 1, "Подготовка к сканированию...");

                HiddenElementsScanResult scanResult = HiddenElementsCollector.Collect(doc, progressWindow);

                progressWindow.Close();
                progressWindow = null;

                if (scanResult.TotalPlansCount == 0)
                {
                    TaskDialog.Show(PluginName, "На листах не найдено размещённых планов.");
                    return Result.Succeeded;
                }

                HiddenElementsFilterWindow window = new HiddenElementsFilterWindow(doc, scanResult);
                bool? dialogResult = window.ShowDialog();

                if (dialogResult != true)
                    return Result.Succeeded;

                WriteResult writeResult = HiddenElementsCollector.WriteFilterParameter(
                    doc,
                    scanResult.HiddenElementViews,
                    FilterParameterName);

                ViewFilterResult filterResult = HiddenElementsCollector.ApplyViewFiltersToViews(
                    doc,
                    scanResult.Sheets,
                    FilterParameterName,
                    FilterNamePrefix);

                UnhideResult unhideResult = HiddenElementsCollector.UnhideHiddenElements(
                    doc,
                    scanResult.Sheets,
                    window.UnhideAllElements,
                    FilterParameterName);

                string resultText = BuildResultText(
                    scanResult,
                    writeResult,
                    filterResult,
                    unhideResult,
                    window.UnhideAllElements);

                TaskDialog.Show(PluginName, resultText);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                if (progressWindow != null)
                {
                    try
                    {
                        progressWindow.Close();
                    }
                    catch
                    {
                    }
                }

                message = ex.ToString();
                return Result.Failed;
            }
        }

        private static string BuildResultText(
            HiddenElementsScanResult scanResult,
            WriteResult writeResult,
            ViewFilterResult filterResult,
            UnhideResult unhideResult,
            bool unhideAllElements)
        {
            StringBuilder sb = new StringBuilder();

            sb.AppendLine("Готово.");
            sb.AppendLine();

            sb.AppendLine("Планов со скрытыми элементами: " + scanResult.TotalPlansWithHiddenElements);
            sb.AppendLine("Уникальных скрытых элементов: " + scanResult.TotalUniqueHiddenElements + " из " + scanResult.TotalHiddenOccurrences);
            sb.AppendLine();

            sb.AppendLine("Запись в KPLN_Фильтрация: " + writeResult.WrittenCount + " из " + scanResult.TotalUniqueHiddenElements);

            if (writeResult.SkippedCount > 0)
            {
                sb.AppendLine("Причины пропуска:");
                AppendNonZeroLine(sb, "NULL", writeResult.NullElementCount);
                AppendNonZeroLine(sb, "Нет параметра", writeResult.NoParameterCount);
                AppendNonZeroLine(sb, "Параметр ReadOnly", writeResult.ReadOnlyCount);
                AppendNonZeroLine(sb, "Тип параметра не строковый", writeResult.WrongStorageTypeCount);
                AppendNonZeroLine(sb, "Ошибка Set()", writeResult.SetFailedCount);
                sb.AppendLine();
            }

            sb.AppendLine("Фильтры в переопределении видимости/графики: " + filterResult.AppliedViewsCount + " из " + filterResult.TargetViewsCount);
            sb.AppendLine("Элементов скрывается фильтрами: " + filterResult.FilteredElementsCount + " из " + filterResult.TargetElementsCount);

            if (filterResult.TemplatesCheckedCount > 0)
            {
                AppendNonZeroLine(sb, "Шаблонов проверено", filterResult.TemplatesCheckedCount);
                AppendNonZeroLine(sb, "Шаблонов, где снят контроль Filters", filterResult.TemplatesReleasedCount);
                AppendNonZeroLine(sb, "Шаблонов, где Filters уже не контролировались", filterResult.TemplatesAlreadyReleasedCount);
                AppendNonZeroLine(sb, "Ошибка изменения шаблона", filterResult.FailedTemplatesCount);
            }

            if (filterResult.CreatedFiltersCount > 0 ||
                filterResult.UpdatedFiltersCount > 0 ||
                filterResult.ViewsWithoutCategoriesCount > 0 ||
                filterResult.FailedViewsCount > 0)
            {
                AppendNonZeroLine(sb, "Создано фильтров", filterResult.CreatedFiltersCount);
                AppendNonZeroLine(sb, "Обновлено фильтров", filterResult.UpdatedFiltersCount);
                AppendNonZeroLine(sb, "Вид без категорий для фильтра", filterResult.ViewsWithoutCategoriesCount);
                AppendNonZeroLine(sb, "Ошибка применения фильтра", filterResult.FailedViewsCount);
                sb.AppendLine();
            }
            else
            {
                sb.AppendLine();
            }

            sb.AppendLine("Обработано планов: " + unhideResult.ProcessedPlansCount + " из " + unhideResult.TargetPlansCount);
            sb.AppendLine("Отображение элементов на плане: " + unhideResult.UnhiddenElementsCount + " из " + unhideResult.TargetElementsCount);
            sb.AppendLine();

            int totalUnhideIssues =
                unhideResult.MissingViewsCount +
                unhideResult.SkippedPlansWithoutTargetElementsCount +
                unhideResult.NoParameterElementsCount +
                unhideResult.NotCurrentlyHiddenElementsCount +
                unhideResult.FailedElementsCount;

            if (totalUnhideIssues > 0)
            {
                sb.AppendLine("Пропуски:");
                AppendNonZeroLine(sb, "Вид не найден", unhideResult.MissingViewsCount);
                AppendNonZeroLine(sb, "План без подходящих элементов", unhideResult.SkippedPlansWithoutTargetElementsCount);

                if (!unhideAllElements)
                    AppendNonZeroLine(sb, "Элемент не отмечен в KPLN_Фильтрация для этого вида", unhideResult.NoParameterElementsCount);

                AppendNonZeroLine(sb, "Элемент уже не скрыт", unhideResult.NotCurrentlyHiddenElementsCount);
                AppendNonZeroLine(sb, "Ошибка отображения элемента", unhideResult.FailedElementsCount);
            }

            return sb.ToString().TrimEnd();
        }

        private static void AppendNonZeroLine(StringBuilder sb, string text, int count)
        {
            if (count > 0)
                sb.AppendLine("- " + text + ": " + count);
        }
    }

    internal static class HiddenElementsCollector
    {
        public static HiddenElementsScanResult Collect(Document doc, ScanProgressWindow progressWindow)
        {
            HiddenElementsScanResult result = new HiddenElementsScanResult();

            List<ViewSheet> sheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .OrderBy(x => x.SheetNumber)
                .ThenBy(x => x.Name)
                .ToList();

            List<Element> candidateElements = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .ToElements()
                .Where(e => e != null && !(e is View) && !(e is ViewSheet))
                .ToList();

            Dictionary<ElementId, List<ViewPlan>> plansBySheetId =
                new Dictionary<ElementId, List<ViewPlan>>(new ElementIdEqualityComparer());

            List<ViewPlan> allPlans = new List<ViewPlan>();

            foreach (ViewSheet sheet in sheets)
            {
                List<ViewPlan> placedPlans = sheet
                    .GetAllPlacedViews()
                    .Select(id => doc.GetElement(id) as View)
                    .Where(v => v != null && !v.IsTemplate)
                    .OfType<ViewPlan>()
                    .ToList();

                if (placedPlans.Count > 0)
                {
                    plansBySheetId[sheet.Id] = placedPlans;
                    allPlans.AddRange(placedPlans);
                }
            }

            result.TotalPlansCount = allPlans.Count;

            int processedPlans = 0;

            foreach (ViewSheet sheet in sheets)
            {
                List<ViewPlan> placedPlans;
                if (!plansBySheetId.TryGetValue(sheet.Id, out placedPlans))
                    continue;

                if (placedPlans == null || placedPlans.Count == 0)
                    continue;

                SheetHiddenInfo sheetInfo = new SheetHiddenInfo
                {
                    SheetId = sheet.Id,
                    SheetNumber = sheet.SheetNumber,
                    SheetName = sheet.Name
                };

                HashSet<ElementId> uniqueHiddenOnSheet = new HashSet<ElementId>(new ElementIdEqualityComparer());
                int hiddenOccurrencesOnSheet = 0;
                int plansWithHiddenElementsOnSheet = 0;

                foreach (ViewPlan plan in placedPlans)
                {
                    processedPlans++;

                    if (progressWindow != null)
                    {
                        progressWindow.UpdateProgress(
                            processedPlans,
                            result.TotalPlansCount,
                            "Сканирование вида: " + plan.Name);
                        DoEvents();
                    }

                    List<ElementId> hiddenIds = GetPermanentlyHiddenElementIds(plan, candidateElements)
                        .OrderBy(x => IDHelper.ElIdInt(x))
                        .ToList();

                    PlanHiddenInfo planInfo = new PlanHiddenInfo
                    {
                        ViewId = plan.Id,
                        ViewName = plan.Name,
                        HiddenCount = hiddenIds.Count,
                        HiddenElementIds = hiddenIds
                    };

                    sheetInfo.Plans.Add(planInfo);

                    hiddenOccurrencesOnSheet += hiddenIds.Count;
                    result.TotalHiddenOccurrences += hiddenIds.Count;

                    if (hiddenIds.Count > 0)
                    {
                        plansWithHiddenElementsOnSheet++;
                        result.TotalPlansWithHiddenElements++;
                    }

                    foreach (ElementId hiddenId in hiddenIds)
                    {
                        uniqueHiddenOnSheet.Add(hiddenId);

                        HashSet<int> viewIds;
                        if (!result.HiddenElementViews.TryGetValue(hiddenId, out viewIds))
                        {
                            viewIds = new HashSet<int>();
                            result.HiddenElementViews.Add(hiddenId, viewIds);
                        }

                        viewIds.Add(IDHelper.ElIdInt(plan.Id));
                    }
                }

                sheetInfo.PlansCount = sheetInfo.Plans.Count;
                sheetInfo.PlansWithHiddenElementsCount = plansWithHiddenElementsOnSheet;
                sheetInfo.HiddenOccurrencesCount = hiddenOccurrencesOnSheet;
                sheetInfo.HiddenUniqueElementsCount = uniqueHiddenOnSheet.Count;

                result.Sheets.Add(sheetInfo);
            }

            if (progressWindow != null)
            {
                progressWindow.UpdateProgress(result.TotalPlansCount, result.TotalPlansCount, "Сканирование завершено");
                DoEvents();
            }

            return result;
        }

        private static List<ElementId> GetPermanentlyHiddenElementIds(View view, List<Element> candidateElements)
        {
            List<ElementId> result = new List<ElementId>();

            foreach (Element e in candidateElements)
            {
                if (e == null)
                    continue;

                if (IDHelper.ElIdInt(e.Id) == IDHelper.ElIdInt(view.Id))
                    continue;

                if (e.ViewSpecific && IDHelper.ElIdInt(e.OwnerViewId) != IDHelper.ElIdInt(view.Id))
                    continue;

                try
                {
                    if (!e.CanBeHidden(view))
                        continue;
                }
                catch
                {
                    continue;
                }

                try
                {
                    if (e.IsHidden(view))
                        result.Add(e.Id);
                }
                catch
                {
                }
            }

            return result;
        }

        public static WriteResult WriteFilterParameter(
            Document doc,
            Dictionary<ElementId, HashSet<int>> hiddenElementViews,
            string parameterName)
        {
            WriteResult result = new WriteResult();

            using (Transaction t = new Transaction(doc, "Запись KPLN_Фильтрация"))
            {
                t.Start();

                foreach (KeyValuePair<ElementId, HashSet<int>> pair in hiddenElementViews)
                {
                    Element element = doc.GetElement(pair.Key);
                    if (element == null)
                    {
                        result.SkippedCount++;
                        result.NullElementCount++;
                        continue;
                    }

                    Parameter parameter = element.LookupParameter(parameterName);
                    if (parameter == null)
                    {
                        result.SkippedCount++;
                        result.NoParameterCount++;
                        continue;
                    }

                    if (parameter.IsReadOnly)
                    {
                        result.SkippedCount++;
                        result.ReadOnlyCount++;
                        continue;
                    }

                    if (parameter.StorageType != StorageType.String)
                    {
                        result.SkippedCount++;
                        result.WrongStorageTypeCount++;
                        continue;
                    }

                    List<int> sortedViewIds = pair.Value
                        .Distinct()
                        .OrderBy(x => x)
                        .ToList();

                    string value = ";" + string.Join(";", sortedViewIds) + ";";

                    try
                    {
                        bool ok = parameter.Set(value);
                        if (ok)
                        {
                            result.WrittenCount++;
                        }
                        else
                        {
                            result.SkippedCount++;
                            result.SetFailedCount++;
                        }
                    }
                    catch
                    {
                        result.SkippedCount++;
                        result.SetFailedCount++;
                    }
                }

                t.Commit();
            }

            return result;
        }

        public static ViewFilterResult ApplyViewFiltersToViews(
            Document doc,
            List<SheetHiddenInfo> sheets,
            string parameterName,
            string filterNamePrefix)
        {
            ViewFilterResult result = new ViewFilterResult();

            Dictionary<string, ParameterFilterElement> existingFilters =
                new FilteredElementCollector(doc)
                    .OfClass(typeof(ParameterFilterElement))
                    .Cast<ParameterFilterElement>()
                    .GroupBy(x => x.Name)
                    .Select(x => x.First())
                    .ToDictionary(x => x.Name, x => x);

            Dictionary<int, bool> templatePrepared = new Dictionary<int, bool>();

            using (Transaction t = new Transaction(doc, "Применение фильтров к видам"))
            {
                t.Start();

                foreach (SheetHiddenInfo sheet in sheets)
                {
                    if (sheet == null || sheet.Plans == null)
                        continue;

                    foreach (PlanHiddenInfo planInfo in sheet.Plans)
                    {
                        if (planInfo == null || planInfo.HiddenElementIds == null || planInfo.HiddenElementIds.Count == 0)
                            continue;

                        View view = doc.GetElement(planInfo.ViewId) as View;
                        if (view == null)
                        {
                            result.FailedViewsCount++;
                            continue;
                        }

                        if (!PrepareTemplateForViewFilters(doc, view, templatePrepared, result))
                        {
                            result.FailedViewsCount++;
                            continue;
                        }

                        List<Element> targetElements = new List<Element>();
                        HashSet<ElementId> categoryIds = new HashSet<ElementId>(new ElementIdEqualityComparer());
                        ElementId parameterId = null;

                        foreach (ElementId id in planInfo.HiddenElementIds.Distinct(new ElementIdLinqComparer()))
                        {
                            Element element = doc.GetElement(id);
                            if (element == null)
                                continue;

                            Parameter parameter = element.LookupParameter(parameterName);
                            if (parameter == null)
                                continue;

                            if (parameter.StorageType != StorageType.String)
                                continue;

                            string paramValue = parameter.AsString();
                            if (!ContainsViewToken(paramValue, IDHelper.ElIdInt(planInfo.ViewId)))
                                continue;

                            if (parameterId == null)
                                parameterId = parameter.Id;

                            if (element.Category != null)
                                categoryIds.Add(element.Category.Id);

                            targetElements.Add(element);
                        }

                        if (targetElements.Count == 0)
                            continue;

                        result.TargetViewsCount++;
                        result.TargetElementsCount += targetElements.Count;

                        List<ElementId> filterCategoryIds = categoryIds.ToList();

                        try
                        {
                            ParameterFilterUtilities.RemoveUnfilterableCategories(filterCategoryIds);
                        }
                        catch
                        {
                        }

                        if (filterCategoryIds.Count == 0 || parameterId == null)
                        {
                            result.ViewsWithoutCategoriesCount++;
                            continue;
                        }

                        try
                        {
                            string token = ";" + IDHelper.ElIdInt(planInfo.ViewId) + ";";

                            FilterRule rule = IDHelper.CreateContainsRule(parameterId, token);

                            ElementParameterFilter elementFilter = new ElementParameterFilter(rule);

                            string filterName = filterNamePrefix + IDHelper.ElIdInt(planInfo.ViewId);
                            ParameterFilterElement parameterFilter;

                            if (existingFilters.TryGetValue(filterName, out parameterFilter))
                            {
                                parameterFilter.SetCategories(filterCategoryIds);
                                parameterFilter.SetElementFilter(elementFilter);
                                result.UpdatedFiltersCount++;
                            }
                            else
                            {
                                parameterFilter = ParameterFilterElement.Create(doc, filterName, filterCategoryIds);
                                parameterFilter.SetElementFilter(elementFilter);
                                existingFilters[filterName] = parameterFilter;
                                result.CreatedFiltersCount++;
                            }

                            ICollection<ElementId> viewFilters = view.GetFilters();
                            if (!viewFilters.Contains(parameterFilter.Id))
                                view.AddFilter(parameterFilter.Id);

                            view.SetFilterVisibility(parameterFilter.Id, false);

                            result.AppliedViewsCount++;
                            result.FilteredElementsCount += targetElements.Count;
                        }
                        catch
                        {
                            result.FailedViewsCount++;
                        }
                    }
                }

                t.Commit();
            }

            return result;
        }

        private static bool PrepareTemplateForViewFilters(
            Document doc,
            View view,
            Dictionary<int, bool> templatePrepared,
            ViewFilterResult result)
        {
            if (doc == null || view == null)
                return false;

            if (view.ViewTemplateId == null ||
                IDHelper.ElIdInt(view.ViewTemplateId) == IDHelper.ElIdInt(ElementId.InvalidElementId))
            {
                return true;
            }

            int templateId = IDHelper.ElIdInt(view.ViewTemplateId);

            bool isReady;
            if (templatePrepared.TryGetValue(templateId, out isReady))
                return isReady;

            View templateView = doc.GetElement(view.ViewTemplateId) as View;
            if (templateView == null || !templateView.IsTemplate)
            {
                templatePrepared[templateId] = false;
                result.FailedTemplatesCount++;
                return false;
            }

            result.TemplatesCheckedCount++;

            try
            {
                ElementId filtersParameterId = new ElementId(BuiltInParameter.VIS_GRAPHICS_FILTERS);
                ICollection<ElementId> nonControlledIds = templateView.GetNonControlledTemplateParameterIds();

                bool alreadyReleased = nonControlledIds.Any(x => x != null && IDHelper.ElIdInt(x) == IDHelper.ElIdInt(filtersParameterId));

                if (alreadyReleased)
                {
                    result.TemplatesAlreadyReleasedCount++;
                    templatePrepared[templateId] = true;
                    return true;
                }

                List<ElementId> newNonControlledIds = nonControlledIds.ToList();
                newNonControlledIds.Add(filtersParameterId);

                templateView.SetNonControlledTemplateParameterIds(newNonControlledIds);

                result.TemplatesReleasedCount++;
                templatePrepared[templateId] = true;
                return true;
            }
            catch
            {
                result.FailedTemplatesCount++;
                templatePrepared[templateId] = false;
                return false;
            }
        }

        public static UnhideResult UnhideHiddenElements(
            Document doc,
            List<SheetHiddenInfo> sheets,
            bool unhideAllElements,
            string parameterName)
        {
            UnhideResult result = new UnhideResult();

            using (Transaction t = new Transaction(doc, "Показ скрытых элементов на планах"))
            {
                t.Start();

                foreach (SheetHiddenInfo sheet in sheets)
                {
                    if (sheet == null || sheet.Plans == null)
                        continue;

                    foreach (PlanHiddenInfo planInfo in sheet.Plans)
                    {
                        if (planInfo == null)
                            continue;

                        if (planInfo.HiddenElementIds == null || planInfo.HiddenElementIds.Count == 0)
                            continue;

                        List<ElementId> targetIdsByParameter = new List<ElementId>();

                        foreach (ElementId id in planInfo.HiddenElementIds)
                        {
                            Element element = doc.GetElement(id);
                            if (element == null)
                                continue;

                            if (!unhideAllElements)
                            {
                                if (!ElementContainsViewToken(element, parameterName, IDHelper.ElIdInt(planInfo.ViewId)))
                                {
                                    result.NoParameterElementsCount++;
                                    continue;
                                }
                            }

                            targetIdsByParameter.Add(id);
                        }

                        if (targetIdsByParameter.Count == 0)
                        {
                            result.SkippedPlansWithoutTargetElementsCount++;
                            continue;
                        }

                        result.TargetPlansCount++;
                        result.TargetElementsCount += targetIdsByParameter.Count;

                        View view = doc.GetElement(planInfo.ViewId) as View;
                        if (view == null)
                        {
                            result.MissingViewsCount++;
                            continue;
                        }

                        List<ElementId> currentlyHiddenIds = new List<ElementId>();

                        foreach (ElementId id in targetIdsByParameter)
                        {
                            Element element = doc.GetElement(id);
                            if (element == null)
                                continue;

                            try
                            {
                                if (!element.CanBeHidden(view))
                                    continue;
                            }
                            catch
                            {
                                continue;
                            }

                            try
                            {
                                if (element.IsHidden(view))
                                    currentlyHiddenIds.Add(id);
                                else
                                    result.NotCurrentlyHiddenElementsCount++;
                            }
                            catch
                            {
                            }
                        }

                        result.ProcessedPlansCount++;

                        if (currentlyHiddenIds.Count == 0)
                            continue;

                        try
                        {
                            view.UnhideElements(currentlyHiddenIds);
                            result.UnhiddenPlansCount++;
                            result.UnhiddenElementsCount += currentlyHiddenIds.Count;
                        }
                        catch
                        {
                            int unhiddenOnPlan = 0;
                            int failedOnPlan = 0;

                            foreach (ElementId id in currentlyHiddenIds)
                            {
                                try
                                {
                                    view.UnhideElements(new List<ElementId> { id });
                                    unhiddenOnPlan++;
                                }
                                catch
                                {
                                    failedOnPlan++;
                                }
                            }

                            if (unhiddenOnPlan > 0)
                                result.UnhiddenPlansCount++;

                            result.UnhiddenElementsCount += unhiddenOnPlan;
                            result.FailedElementsCount += failedOnPlan;
                        }
                    }
                }

                t.Commit();
            }

            return result;
        }

        private static bool ElementContainsViewToken(Element element, string parameterName, int viewId)
        {
            if (element == null)
                return false;

            Parameter parameter = element.LookupParameter(parameterName);
            if (parameter == null || parameter.StorageType != StorageType.String)
                return false;

            return ContainsViewToken(parameter.AsString(), viewId);
        }

        private static bool ContainsViewToken(string value, int viewId)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;

            string token = ";" + viewId + ";";
            return value.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static void DoEvents()
        {
            DispatcherFrame frame = new DispatcherFrame();
            Dispatcher.CurrentDispatcher.BeginInvoke(
                DispatcherPriority.Background,
                new DispatcherOperationCallback(ExitFrame),
                frame);
            Dispatcher.PushFrame(frame);
        }

        private static object ExitFrame(object frame)
        {
            DispatcherFrame dispatcherFrame = frame as DispatcherFrame;
            if (dispatcherFrame != null)
                dispatcherFrame.Continue = false;

            return null;
        }
    }

    public class HiddenElementsScanResult
    {
        public HiddenElementsScanResult()
        {
            Sheets = new List<SheetHiddenInfo>();
            HiddenElementViews = new Dictionary<ElementId, HashSet<int>>(new ElementIdEqualityComparer());
        }

        public List<SheetHiddenInfo> Sheets { get; private set; }
        public Dictionary<ElementId, HashSet<int>> HiddenElementViews { get; private set; }

        public int TotalPlansCount { get; set; }
        public int TotalPlansWithHiddenElements { get; set; }
        public int TotalHiddenOccurrences { get; set; }

        public int TotalUniqueHiddenElements
        {
            get { return HiddenElementViews.Count; }
        }
    }

    public class SheetHiddenInfo
    {
        public SheetHiddenInfo()
        {
            Plans = new List<PlanHiddenInfo>();
        }

        public ElementId SheetId { get; set; }
        public string SheetNumber { get; set; }
        public string SheetName { get; set; }

        public int PlansCount { get; set; }
        public int PlansWithHiddenElementsCount { get; set; }
        public int HiddenOccurrencesCount { get; set; }
        public int HiddenUniqueElementsCount { get; set; }

        public List<PlanHiddenInfo> Plans { get; set; }

        public string SheetDisplayName
        {
            get { return SheetNumber + " - " + SheetName; }
        }
    }

    public class PlanHiddenInfo
    {
        public PlanHiddenInfo()
        {
            HiddenElementIds = new List<ElementId>();
        }

        public ElementId ViewId { get; set; }
        public string ViewName { get; set; }
        public int HiddenCount { get; set; }
        public List<ElementId> HiddenElementIds { get; set; }
    }

    public class WriteResult
    {
        public int WrittenCount { get; set; }
        public int SkippedCount { get; set; }

        public int NullElementCount { get; set; }
        public int NoParameterCount { get; set; }
        public int ReadOnlyCount { get; set; }
        public int WrongStorageTypeCount { get; set; }
        public int SetFailedCount { get; set; }
    }

    public class ViewFilterResult
    {
        public int TargetViewsCount { get; set; }
        public int AppliedViewsCount { get; set; }

        public int TargetElementsCount { get; set; }
        public int FilteredElementsCount { get; set; }

        public int CreatedFiltersCount { get; set; }
        public int UpdatedFiltersCount { get; set; }
        public int ViewsWithoutCategoriesCount { get; set; }
        public int FailedViewsCount { get; set; }

        public int TemplatesCheckedCount { get; set; }
        public int TemplatesReleasedCount { get; set; }
        public int TemplatesAlreadyReleasedCount { get; set; }
        public int FailedTemplatesCount { get; set; }
    }

    public class UnhideResult
    {
        public int TargetPlansCount { get; set; }
        public int ProcessedPlansCount { get; set; }
        public int UnhiddenPlansCount { get; set; }

        public int TargetElementsCount { get; set; }
        public int UnhiddenElementsCount { get; set; }

        public int MissingViewsCount { get; set; }
        public int SkippedPlansWithoutTargetElementsCount { get; set; }
        public int NoParameterElementsCount { get; set; }
        public int NotCurrentlyHiddenElementsCount { get; set; }
        public int FailedElementsCount { get; set; }
    }

    internal class ElementIdEqualityComparer : IEqualityComparer<ElementId>
    {
        public bool Equals(ElementId x, ElementId y)
        {
            if (ReferenceEquals(x, y))
                return true;

            if (ReferenceEquals(x, null) || ReferenceEquals(y, null))
                return false;

            return IDHelper.ElIdInt(x) == IDHelper.ElIdInt(y);
        }

        public int GetHashCode(ElementId obj)
        {
            return obj == null ? 0 : IDHelper.ElIdInt(obj).GetHashCode();
        }
    }

    internal class ElementIdLinqComparer : IEqualityComparer<ElementId>
    {
        public bool Equals(ElementId x, ElementId y)
        {
            if (ReferenceEquals(x, y))
                return true;

            if (ReferenceEquals(x, null) || ReferenceEquals(y, null))
                return false;

            return IDHelper.ElIdInt(x) == IDHelper.ElIdInt(y);
        }

        public int GetHashCode(ElementId obj)
        {
            return obj == null ? 0 : IDHelper.ElIdInt(obj).GetHashCode();
        }
    }
}