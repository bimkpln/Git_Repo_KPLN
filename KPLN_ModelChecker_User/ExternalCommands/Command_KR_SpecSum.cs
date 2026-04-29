using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace KPLN_ModelChecker_User
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class Command_KR_SpecSum : IExternalCommand
    {
        internal const string PluginName = "Сумма спецификаций";

        private const string TargetParamName = "Орг.КомплектЧертежей";
        private const string DefaultTargetParamValue = "КЖ";

        private static readonly string[] TargetFieldNames =
        {
            "МассаОбщ",
            "МассаВсего"
        };

        private static readonly string[] GeneralDataFieldNames =
        {
            "МассаАрмОбщ",
            "МассаЗДОбщ"
        };

        private static readonly string[] GeneralDataNameParts =
        {
            "общие данные"
        };

        private static readonly string[] SteelScheduleNameParts =
        {
            "врс",
            "ведомость расхода стали",
            "расхода стали"
        };

        private static readonly IComparer<string> NaturalComparer = new NaturalStringComparer();

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            SheetFilterWindow sheetFilterWindow = new SheetFilterWindow(
                TargetParamName,
                DefaultTargetParamValue);

            bool? sheetFilterResult = sheetFilterWindow.ShowDialog();
            if (sheetFilterResult != true)
            {
                return Result.Cancelled;
            }

            string targetParamValue = sheetFilterWindow.FilterValue;

            List<ViewSheet> sheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .ToList();

            SheetFilterResult filterResult = CollectSheetsByParamContains(
                doc,
                TargetParamName,
                targetParamValue);

            if (filterResult.Sheets.Count == 0)
            {
                TaskDialog.Show(
                    PluginName,
                    string.Format(
                        CultureInfo.InvariantCulture,
                        "Не найдено листов, у которых параметр '{0}' содержит '{1}'.",
                        TargetParamName,
                        targetParamValue));
                return Result.Cancelled;
            }

            ViewSheet generalDataSheet = sheets
                .OrderBy(x => x.SheetNumber, NaturalComparer)
                .ThenBy(x => x.Name, NaturalComparer)
                .FirstOrDefault(x => ContainsAnyPart(x.Name, GeneralDataNameParts));

            HashSet<int> relevantSheetIds = new HashSet<int>(
                filterResult.Sheets.Select(x => IDHelper.ElIdInt(x.Id)));

            if (generalDataSheet != null)
            {
                relevantSheetIds.Add(IDHelper.ElIdInt(generalDataSheet.Id));
            }

            Dictionary<int, List<ScheduleSheetInstance>> scheduleInstancesBySheet =
                CollectScheduleInstancesBySheet(doc, relevantSheetIds);

            Dictionary<int, ViewSchedule> scheduleCache = new Dictionary<int, ViewSchedule>();
            StringBuilder report = new StringBuilder();
            List<SpecSumReportItem> reportItems = new List<SpecSumReportItem>();
            string reportDateTime = DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss", CultureInfo.InvariantCulture);
            string mainReportTitle = string.Format(
                CultureInfo.InvariantCulture,
                "Отчет по спецификациям армирования ({0})",
                reportDateTime);

            double grandTotalSum = 0.0;
            int grandTotalFoundValues = 0;

            if (filterResult.UsedFallback)
            {
                PrintWarning(
                    reportItems,
                    report,
                    string.Format(
                        "Фильтр по параметру '{0}' недоступен - листы отобраны перебором. Результаты могут быть неполными.",
                        TargetParamName));
            }

            PrintReportTitle(reportItems, report, mainReportTitle);

            foreach (ViewSheet sheet in filterResult.Sheets
                .OrderBy(x => x.SheetNumber, NaturalComparer)
                .ThenBy(x => x.Name, NaturalComparer))
            {
                List<ScheduleSheetInstance> sheetInstances;
                if (!scheduleInstancesBySheet.TryGetValue(IDHelper.ElIdInt(sheet.Id), out sheetInstances))
                {
                    sheetInstances = new List<ScheduleSheetInstance>();
                }

                Dictionary<int, ViewSchedule> matchedSchedules = new Dictionary<int, ViewSchedule>();

                foreach (ScheduleSheetInstance instance in sheetInstances)
                {
                    ViewSchedule schedule = GetCachedSchedule(doc, scheduleCache, instance.ScheduleId);
                    if (schedule == null)
                    {
                        continue;
                    }

                    if (ScheduleHasAnyFieldName(schedule, TargetFieldNames))
                    {
                        int scheduleId = IDHelper.ElIdInt(schedule.Id);
                        if (!matchedSchedules.ContainsKey(scheduleId))
                        {
                            matchedSchedules.Add(scheduleId, schedule);
                        }
                    }
                }

                foreach (int scheduleId in matchedSchedules.Keys.OrderBy(x => x))
                {
                    ViewSchedule schedule = matchedSchedules[scheduleId];
                    ScheduleCalculationResult calculation = CalculateScheduleSum(schedule);

                    if (calculation.FoundValues > 0)
                    {
                        grandTotalSum += calculation.TotalSum;
                        grandTotalFoundValues += calculation.FoundValues;
                        PrintScheduleBlock(
                            reportItems,
                            report,
                            sheet.SheetNumber,
                            sheet.Name,
                            schedule.Name,
                            FormatNumber(calculation.TotalSum),
                            false);
                    }
                    else
                    {
                        PrintScheduleBlock(
                            reportItems,
                            report,
                            sheet.SheetNumber,
                            sheet.Name,
                            schedule.Name,
                            "нет данных",
                            true);
                    }
                }
            }

            if (grandTotalFoundValues > 0)
            {
                PrintTotalLine(reportItems, report, "Итоговая сумма:", FormatNumber(grandTotalSum), true);
            }

            List<ViewSchedule> generalDataSchedules = new List<ViewSchedule>();
            double generalDataTotalSum = 0.0;

            if (generalDataSheet != null)
            {
                List<ScheduleSheetInstance> sheetInstances;
                if (!scheduleInstancesBySheet.TryGetValue(IDHelper.ElIdInt(generalDataSheet.Id), out sheetInstances))
                {
                    sheetInstances = new List<ScheduleSheetInstance>();
                }

                HashSet<int> generalDataScheduleIds = new HashSet<int>();

                foreach (ScheduleSheetInstance instance in sheetInstances)
                {
                    ViewSchedule schedule = GetCachedSchedule(doc, scheduleCache, instance.ScheduleId);
                    if (schedule == null)
                    {
                        continue;
                    }

                    int scheduleId = IDHelper.ElIdInt(schedule.Id);
                    if (ContainsAnyPart(schedule.Name, SteelScheduleNameParts)
                        && ScheduleHasAnyFieldName(schedule, GeneralDataFieldNames)
                        && !generalDataScheduleIds.Contains(scheduleId))
                    {
                        generalDataScheduleIds.Add(scheduleId);
                        generalDataSchedules.Add(schedule);

                        if (generalDataSchedules.Count >= 2)
                        {
                            break;
                        }
                    }
                }
            }

            if (generalDataSheet != null && generalDataSchedules.Count > 0)
            {
                PrintReportTitle(reportItems, report, "Общие данные");

                foreach (ViewSchedule generalDataSchedule in generalDataSchedules)
                {
                    ScheduleCalculationResult calculation = AnalyzeGeneralDataSchedule(generalDataSchedule);

                    if (calculation.FoundValues > 0)
                    {
                        generalDataTotalSum += calculation.TotalSum;
                        PrintScheduleBlock(
                            reportItems,
                            report,
                            generalDataSheet.SheetNumber,
                            generalDataSheet.Name,
                            generalDataSchedule.Name,
                            FormatNumber(calculation.TotalSum),
                            false);
                    }
                    else
                    {
                        PrintScheduleBlock(
                            reportItems,
                            report,
                            generalDataSheet.SheetNumber,
                            generalDataSheet.Name,
                            generalDataSchedule.Name,
                            "нет данных",
                            true);
                    }
                }

                PrintTotalLine(reportItems, report, "Итог по общим данным:", FormatNumber(generalDataTotalSum), true);
            }

            if (grandTotalFoundValues > 0 && generalDataTotalSum > 0)
            {
                double differenceValue = Math.Abs(grandTotalSum - generalDataTotalSum);
                double differencePercent = differenceValue / generalDataTotalSum * 100.0;

                PrintReportTitle(reportItems, report, "Сравнение итогов");
                PrintTotalLine(reportItems, report, "Разница:", FormatNumber(differenceValue), false);
                PrintTotalLine(
                    reportItems,
                    report,
                    "Отклонение от итога по общим данным:",
                    FormatPercent(differencePercent),
                    false);
            }

            SpecSumReportWindow reportWindow = new SpecSumReportWindow(PluginName, reportItems, report.ToString());
            reportWindow.ShowDialog();

            return Result.Succeeded;
        }

        private static SheetFilterResult CollectSheetsByParamContains(
            Document doc,
            string paramName,
            string paramValue)
        {
            List<ViewSheet> allSheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .ToList();

            ElementId sampleParameterId = null;
            foreach (ViewSheet sheet in allSheets)
            {
                Parameter parameter = sheet.LookupParameter(paramName);
                if (parameter != null)
                {
                    sampleParameterId = parameter.Id;
                    break;
                }
            }

            if (sampleParameterId != null)
            {
                try
                {
                    FilterRule rule = IDHelper.CreateContainsRule(sampleParameterId, paramValue);
                    ElementParameterFilter paramFilter = new ElementParameterFilter(rule);

                    List<ViewSheet> filteredSheets = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewSheet))
                        .WherePasses(paramFilter)
                        .Cast<ViewSheet>()
                        .ToList();

                    return new SheetFilterResult(filteredSheets, false);
                }
                catch
                {
                }
            }

            List<ViewSheet> matchedSheets = new List<ViewSheet>();
            foreach (ViewSheet sheet in allSheets)
            {
                string value = GetStringParameterValue(sheet, paramName);
                if (Contains(value, paramValue))
                {
                    matchedSheets.Add(sheet);
                }
            }

            return new SheetFilterResult(matchedSheets, true);
        }

        private static Dictionary<int, List<ScheduleSheetInstance>> CollectScheduleInstancesBySheet(
            Document doc,
            HashSet<int> relevantSheetIds)
        {
            Dictionary<int, List<ScheduleSheetInstance>> result =
                new Dictionary<int, List<ScheduleSheetInstance>>();

            IEnumerable<ScheduleSheetInstance> allScheduleInstances = new FilteredElementCollector(doc)
                .OfClass(typeof(ScheduleSheetInstance))
                .Cast<ScheduleSheetInstance>();

            foreach (ScheduleSheetInstance instance in allScheduleInstances)
            {
                int ownerViewId = IDHelper.ElIdInt(instance.OwnerViewId);
                if (!relevantSheetIds.Contains(ownerViewId))
                {
                    continue;
                }

                List<ScheduleSheetInstance> instances;
                if (!result.TryGetValue(ownerViewId, out instances))
                {
                    instances = new List<ScheduleSheetInstance>();
                    result.Add(ownerViewId, instances);
                }

                instances.Add(instance);
            }

            return result;
        }

        private static ViewSchedule GetCachedSchedule(
            Document doc,
            Dictionary<int, ViewSchedule> scheduleCache,
            ElementId scheduleId)
        {
            int id = IDHelper.ElIdInt(scheduleId);

            ViewSchedule schedule;
            if (scheduleCache.TryGetValue(id, out schedule))
            {
                return schedule;
            }

            schedule = doc.GetElement(scheduleId) as ViewSchedule;
            scheduleCache[id] = schedule;
            return schedule;
        }

        private static ScheduleCalculationResult CalculateScheduleSum(ViewSchedule schedule)
        {
            TableData tableData = schedule.GetTableData();
            TableSectionData bodyData = tableData.GetSectionData(SectionType.Body);

            double totalSum = 0.0;
            int foundValues = 0;

            for (int row = bodyData.FirstRowNumber; row <= bodyData.LastRowNumber; row++)
            {
                double? lastNumericValue = null;

                for (int col = bodyData.FirstColumnNumber; col <= bodyData.LastColumnNumber; col++)
                {
                    string cellText = string.Empty;
                    try
                    {
                        cellText = schedule.GetCellText(SectionType.Body, row, col);
                    }
                    catch
                    {
                    }

                    double? number = ParseNumber(cellText);
                    if (number.HasValue)
                    {
                        lastNumericValue = number.Value;
                    }
                }

                if (lastNumericValue.HasValue)
                {
                    totalSum += lastNumericValue.Value;
                    foundValues++;
                }
            }

            return new ScheduleCalculationResult(totalSum, foundValues);
        }

        private static ScheduleCalculationResult AnalyzeGeneralDataSchedule(ViewSchedule schedule)
        {
            TableData tableData = schedule.GetTableData();
            TableSectionData bodyData = tableData.GetSectionData(SectionType.Body);
            Dictionary<string, int> visibleFieldColumns = GetScheduleVisibleFieldColumns(schedule);

            string[] armFallbackFieldNames =
            {
                "ИтогоВр1",
                "ИтогоА240",
                "ИтогоА500",
                "Лист итого",
                "Арм Уголок Общ",
                "Арм Полоса Общ"
            };

            List<int> targetColumns = new List<int>();

            if (visibleFieldColumns.ContainsKey("МассаАрмОбщ"))
            {
                AddVisibleFieldColumn("МассаАрмОбщ", visibleFieldColumns, bodyData, targetColumns);
            }
            else
            {
                foreach (string fieldName in armFallbackFieldNames)
                {
                    AddVisibleFieldColumn(fieldName, visibleFieldColumns, bodyData, targetColumns);
                }
            }

            if (visibleFieldColumns.ContainsKey("МассаЗДОбщ"))
            {
                AddVisibleFieldColumn("МассаЗДОбщ", visibleFieldColumns, bodyData, targetColumns);
            }

            if (targetColumns.Count == 0)
            {
                return new ScheduleCalculationResult(0.0, 0);
            }

            bool useGrandTotalRow = false;
            try
            {
                useGrandTotalRow = schedule.Definition.ShowGrandTotal;
            }
            catch
            {
            }

            if (useGrandTotalRow)
            {
                double grandTotalSum = 0.0;
                bool foundInTotal = false;
                int totalRow = bodyData.LastRowNumber;

                foreach (int col in targetColumns)
                {
                    string cellText = string.Empty;
                    try
                    {
                        cellText = schedule.GetCellText(SectionType.Body, totalRow, col);
                    }
                    catch
                    {
                    }

                    double? number = ParseNumber(cellText);
                    if (number.HasValue)
                    {
                        grandTotalSum += number.Value;
                        foundInTotal = true;
                    }
                }

                if (foundInTotal)
                {
                    return new ScheduleCalculationResult(grandTotalSum, 1);
                }
            }

            double totalSum = 0.0;
            int foundValues = 0;

            for (int row = bodyData.FirstRowNumber; row <= bodyData.LastRowNumber; row++)
            {
                double rowSum = 0.0;
                bool rowHasValue = false;

                foreach (int col in targetColumns)
                {
                    string cellText = string.Empty;
                    try
                    {
                        cellText = schedule.GetCellText(SectionType.Body, row, col);
                    }
                    catch
                    {
                    }

                    double? number = ParseNumber(cellText);
                    if (number.HasValue)
                    {
                        rowSum += number.Value;
                        rowHasValue = true;
                    }
                }

                if (rowHasValue)
                {
                    totalSum += rowSum;
                    foundValues++;
                }
            }

            return new ScheduleCalculationResult(totalSum, foundValues);
        }

        private static Dictionary<string, int> GetScheduleVisibleFieldColumns(ViewSchedule schedule)
        {
            Dictionary<string, int> visibleColumns = new Dictionary<string, int>(StringComparer.Ordinal);

            ScheduleDefinition definition;
            int fieldCount;
            try
            {
                definition = schedule.Definition;
                fieldCount = definition.GetFieldCount();
            }
            catch
            {
                return visibleColumns;
            }

            int visibleColumnIndex = 0;
            for (int index = 0; index < fieldCount; index++)
            {
                ScheduleField field;
                string fieldName;
                bool isHidden;
                try
                {
                    field = definition.GetField(index);
                    fieldName = field.GetName();
                    isHidden = field.IsHidden;
                }
                catch
                {
                    continue;
                }

                if (isHidden)
                {
                    continue;
                }

                visibleColumns[fieldName ?? string.Empty] = visibleColumnIndex;
                visibleColumnIndex++;
            }

            return visibleColumns;
        }

        private static void AddVisibleFieldColumn(
            string fieldName,
            Dictionary<string, int> visibleFieldColumns,
            TableSectionData bodyData,
            List<int> targetColumns)
        {
            int visibleIndex;
            if (!visibleFieldColumns.TryGetValue(fieldName, out visibleIndex))
            {
                return;
            }

            int col = bodyData.FirstColumnNumber + visibleIndex;
            if (col > bodyData.LastColumnNumber)
            {
                return;
            }

            if (!targetColumns.Contains(col))
            {
                targetColumns.Add(col);
            }
        }

        private static bool ScheduleHasFieldName(ViewSchedule schedule, string fieldName)
        {
            ScheduleDefinition definition;
            int fieldCount;
            try
            {
                definition = schedule.Definition;
                fieldCount = definition.GetFieldCount();
            }
            catch
            {
                return false;
            }

            for (int index = 0; index < fieldCount; index++)
            {
                string currentName = null;
                try
                {
                    currentName = definition.GetField(index).GetName();
                }
                catch
                {
                }

                if (!string.IsNullOrEmpty(currentName)
                    && string.Equals(currentName, fieldName ?? string.Empty, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool ScheduleHasAnyFieldName(ViewSchedule schedule, IEnumerable<string> fieldNames)
        {
            return fieldNames.Any(fieldName => ScheduleHasFieldName(schedule, fieldName));
        }

        private static string GetStringParameterValue(Element element, string paramName)
        {
            Parameter parameter = element.LookupParameter(paramName);
            if (parameter == null)
            {
                return string.Empty;
            }

            string value = null;
            try
            {
                value = parameter.AsString();
            }
            catch
            {
            }

            if (!string.IsNullOrEmpty(value))
            {
                return value;
            }

            try
            {
                value = parameter.AsValueString();
            }
            catch
            {
                value = null;
            }

            return value ?? string.Empty;
        }

        private static double? ParseNumber(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return null;
            }

            string cleaned = text
                .Trim()
                .Replace(" ", string.Empty)
                .Replace("\u00a0", string.Empty)
                .Replace(",", ".");

            double result;
            if (double.TryParse(cleaned, NumberStyles.Float, CultureInfo.InvariantCulture, out result))
            {
                return result;
            }

            return null;
        }

        private static string FormatNumber(double value)
        {
            return value.ToString("0.00", CultureInfo.InvariantCulture).Replace(".", ",");
        }

        private static string FormatPercent(double value)
        {
            return value.ToString("0.00", CultureInfo.InvariantCulture).Replace(".", ",") + "%";
        }

        private static bool ContainsAnyPart(string name, IEnumerable<string> parts)
        {
            if (string.IsNullOrEmpty(name))
            {
                return false;
            }

            return parts.Any(part => Contains(name, part));
        }

        private static bool Contains(string source, string value)
        {
            if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(value))
            {
                return false;
            }

            return source.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static void PrintWarning(
            List<SpecSumReportItem> items,
            StringBuilder output,
            string message)
        {
            items.Add(new SpecSumReportItem
            {
                Kind = SpecSumReportItemKind.Warning,
                Message = message
            });

            output.AppendLine("ВНИМАНИЕ: " + message);
            output.AppendLine();
        }

        private static void PrintReportTitle(
            List<SpecSumReportItem> items,
            StringBuilder output,
            string title)
        {
            items.Add(new SpecSumReportItem
            {
                Kind = SpecSumReportItemKind.Title,
                Title = title
            });

            if (output.Length > 0)
            {
                output.AppendLine();
            }

            output.AppendLine(title);
            output.AppendLine(new string('=', title.Length));
            output.AppendLine();
        }

        private static string PrintSheetTitle(string sheetNumber, string sheetName)
        {
            return string.Format(CultureInfo.InvariantCulture, "{0} | {1}", sheetNumber, sheetName);
        }

        private static void PrintScheduleBlock(
            List<SpecSumReportItem> items,
            StringBuilder output,
            string sheetNumber,
            string sheetName,
            string scheduleName,
            string valueText,
            bool muted)
        {
            items.Add(new SpecSumReportItem
            {
                Kind = SpecSumReportItemKind.ScheduleBlock,
                SheetNumber = sheetNumber,
                SheetName = sheetName,
                ScheduleName = scheduleName,
                ValueText = valueText,
                Muted = muted
            });

            output.AppendLine(PrintSheetTitle(sheetNumber, sheetName));
            output.AppendLine("  " + scheduleName);
            output.AppendLine(valueText);
            output.AppendLine();
        }

        private static void PrintTotalLine(
            List<SpecSumReportItem> items,
            StringBuilder output,
            string label,
            string valueText,
            bool accent)
        {
            items.Add(new SpecSumReportItem
            {
                Kind = SpecSumReportItemKind.TotalLine,
                Label = label,
                ValueText = valueText,
                Accent = accent
            });

            output.AppendLine(label + " " + valueText);
            output.AppendLine();
        }

        private sealed class SheetFilterResult
        {
            internal SheetFilterResult(List<ViewSheet> sheets, bool usedFallback)
            {
                Sheets = sheets;
                UsedFallback = usedFallback;
            }

            internal List<ViewSheet> Sheets { get; private set; }
            internal bool UsedFallback { get; private set; }
        }

        private sealed class ScheduleCalculationResult
        {
            internal ScheduleCalculationResult(double totalSum, int foundValues)
            {
                TotalSum = totalSum;
                FoundValues = foundValues;
            }

            internal double TotalSum { get; private set; }
            internal int FoundValues { get; private set; }
        }

        private sealed class NaturalStringComparer : IComparer<string>
        {
            public int Compare(string x, string y)
            {
                x = x ?? string.Empty;
                y = y ?? string.Empty;

                int ix = 0;
                int iy = 0;

                while (ix < x.Length && iy < y.Length)
                {
                    bool xDigit = char.IsDigit(x[ix]);
                    bool yDigit = char.IsDigit(y[iy]);

                    if (xDigit && yDigit)
                    {
                        int xStart = ix;
                        int yStart = iy;

                        while (ix < x.Length && char.IsDigit(x[ix]))
                        {
                            ix++;
                        }

                        while (iy < y.Length && char.IsDigit(y[iy]))
                        {
                            iy++;
                        }

                        string xNumber = x.Substring(xStart, ix - xStart).TrimStart('0');
                        string yNumber = y.Substring(yStart, iy - yStart).TrimStart('0');

                        if (xNumber.Length == 0)
                        {
                            xNumber = "0";
                        }

                        if (yNumber.Length == 0)
                        {
                            yNumber = "0";
                        }

                        int lengthCompare = xNumber.Length.CompareTo(yNumber.Length);
                        if (lengthCompare != 0)
                        {
                            return lengthCompare;
                        }

                        int numberCompare = string.CompareOrdinal(xNumber, yNumber);
                        if (numberCompare != 0)
                        {
                            return numberCompare;
                        }

                        continue;
                    }

                    int charCompare = string.Compare(
                        x[ix].ToString(),
                        y[iy].ToString(),
                        StringComparison.OrdinalIgnoreCase);

                    if (charCompare != 0)
                    {
                        return charCompare;
                    }

                    ix++;
                    iy++;
                }

                return (x.Length - ix).CompareTo(y.Length - iy);
            }
        }
    }
}