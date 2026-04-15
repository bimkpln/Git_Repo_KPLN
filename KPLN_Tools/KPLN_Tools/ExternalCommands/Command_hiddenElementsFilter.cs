using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
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

                UnhideResult unhideResult = HiddenElementsCollector.UnhideHiddenElements(
                    doc,
                    scanResult.Sheets,
                    window.UnhideAllElements,
                    FilterParameterName);

                string resultText = BuildResultText(
                    scanResult,
                    writeResult,
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
                AppendNonZeroLine(sb, "Параметр имет не строчный тип", writeResult.WrongStorageTypeCount);
                AppendNonZeroLine(sb, "Ошибка Set()", writeResult.SetFailedCount);
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
                sb.AppendLine("Ошибки и пропуски:");
                AppendNonZeroLine(sb, "Вид не найден", unhideResult.MissingViewsCount);
                AppendNonZeroLine(sb, "План без подходящих элементов", unhideResult.SkippedPlansWithoutTargetElementsCount);

                if (!unhideAllElements)
                    AppendNonZeroLine(sb, "Элемент без параметра KPLN_Фильтрация", unhideResult.NoParameterElementsCount);

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

            Dictionary<ElementId, List<ViewPlan>> plansBySheetId = new Dictionary<ElementId, List<ViewPlan>>(new ElementIdEqualityComparer());
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
                        .OrderBy(x => x.IntegerValue)
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

                        HashSet<string> viewNames;
                        if (!result.HiddenElementViews.TryGetValue(hiddenId, out viewNames))
                        {
                            viewNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                            result.HiddenElementViews.Add(hiddenId, viewNames);
                        }

                        viewNames.Add(plan.Name);
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

                if (e.Id.IntegerValue == view.Id.IntegerValue)
                    continue;

                if (e.ViewSpecific && e.OwnerViewId.IntegerValue != view.Id.IntegerValue)
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
            Dictionary<ElementId, HashSet<string>> hiddenElementViews,
            string parameterName)
        {
            WriteResult result = new WriteResult();

            using (Transaction t = new Transaction(doc, "Запись KPLN_Фильтрация"))
            {
                t.Start();

                foreach (KeyValuePair<ElementId, HashSet<string>> pair in hiddenElementViews)
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

                    string value = string.Join(
                        ";",
                        pair.Value
                            .Where(x => !string.IsNullOrWhiteSpace(x))
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .OrderBy(x => x));

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
                                Parameter parameter = element.LookupParameter(parameterName);
                                if (parameter == null)
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
            HiddenElementViews = new Dictionary<ElementId, HashSet<string>>(new ElementIdEqualityComparer());
        }

        public List<SheetHiddenInfo> Sheets { get; private set; }
        public Dictionary<ElementId, HashSet<string>> HiddenElementViews { get; private set; }

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

            return x.IntegerValue == y.IntegerValue;
        }

        public int GetHashCode(ElementId obj)
        {
            return obj == null ? 0 : obj.IntegerValue.GetHashCode();
        }
    }
}