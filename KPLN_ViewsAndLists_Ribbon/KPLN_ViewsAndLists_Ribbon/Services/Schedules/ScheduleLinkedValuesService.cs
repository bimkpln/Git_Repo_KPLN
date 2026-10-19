using Autodesk.Revit.DB;
using KPLN_ViewsAndLists_Ribbon.Common.Schedules;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace KPLN_ViewsAndLists_Ribbon.Services.Schedules
{
    /// <summary>Читает строки временной спецификации вместе со связями, отфильтрованными самим Revit.</summary>
    internal static class ScheduleLinkedValuesService
    {
        public static List<ScheduleCopyValue> Read(Document doc, ViewSchedule probe, ScheduleField field, ScheduleParameterChoice parameter)
        {
            var definition = probe.Definition;
            if (definition.HasEmbeddedSchedule)
                throw new InvalidOperationException("Для копирования по значениям выберите спецификацию без вложенной таблицы.");

            // Отдельные строки и служебная колонка количества позволяют отличить данные
            // от заголовков и корректно учесть несколько экземпляров одной связи.
            definition.SetSortGroupFields(new List<ScheduleSortGroupField>());
            definition.IsItemized = true;
            definition.ShowTitle = false;
            definition.ShowHeaders = false;
            definition.ShowGrandTotal = false;
            foreach (var id in definition.GetFieldOrder()) definition.GetField(id).IsHidden = id != field.FieldId;
            // Для текстовых полей Revit может отклонить даже повторное присваивание
            // Standard. Меняем режим только у действительно агрегированного поля.
            if (field.DisplayType != ScheduleFieldDisplayType.Standard)
                field.DisplayType = ScheduleFieldDisplayType.Standard;
            field.ColumnHeading = "__KPLN_VALUE__";
            var count = definition.GetFieldOrder().Select(definition.GetField).FirstOrDefault(f => f.FieldType == ScheduleFieldType.Count)
                ?? definition.AddField(ScheduleFieldType.Count);
            count.IsHidden = false;
            count.ColumnHeading = "__KPLN_COUNT__";
            doc.Regenerate();

            var cells = ReadCells(probe, field.FieldId, count.FieldId);
            if (cells.Count == 0) return new List<ScheduleCopyValue>();

            // Текст читается прямо из таблицы, без обхода всех элементов файлов связей.
            if (parameter.SuggestedKind == "Текст")
                return cells.Select(pair => new ScheduleCopyValue
                {
                    Storage = StorageType.String, Text = pair.Key, Display = pair.Key, ElementCount = pair.Value
                }).OrderBy(value => value.Display, StringComparer.CurrentCultureIgnoreCase).ToList();

            var candidates = CollectCandidates(doc, definition, field);
            var result = new Dictionary<string, ScheduleCopyValue>(StringComparer.OrdinalIgnoreCase);
            var originalFilters = definition.GetFilters().ToList();
            foreach (var cell in cells)
            {
                List<ScheduleCopyValue> matching;
                if (!candidates.TryGetValue(CaptionKey(cell.Key), out matching) || matching.Count == 0)
                    throw new InvalidOperationException("Не удалось сопоставить точное значение «" + cell.Key + "» параметра «"
                        + parameter.Name + "». Для ссылок на элементы требуется значение, доступное для фильтра в основном файле.");

                if (matching.Count == 1 && matching[0].Storage != StorageType.ElementId)
                {
                    AddResult(result, matching[0], cell.Key, cell.Value);
                    continue;
                }

                // Округлённые подписи могут совпадать у разных чисел. Проверяем исходные
                // значения фильтром Revit, а не превращаем подпись ячейки обратно в число.
                int matchedRows = 0;
                foreach (var value in matching)
                {
                    var filters = originalFilters.ToList();
                    filters.Add(ScheduleCopyService.CreateEquality(field.FieldId, value));
                    definition.SetFilters(filters);
                    doc.Regenerate();
                    int rows;
                    if (!ReadCells(probe, field.FieldId, count.FieldId).TryGetValue(cell.Key, out rows) || rows == 0) continue;
                    AddResult(result, value, cell.Key, rows);
                    matchedRows += rows;
                }
                if (matchedRows != cell.Value)
                    throw new InvalidOperationException("Revit не смог однозначно разделить значение «" + cell.Key
                        + "» фильтром «Равно». Увеличьте точность отображения параметра или выберите другой параметр копирования.");
            }
            return result.Values.OrderBy(value => value.Storage == StorageType.Double ? value.Number
                : value.Storage == StorageType.Integer ? value.Integer : 0)
                .ThenBy(value => value.Display, StringComparer.CurrentCultureIgnoreCase).ToList();
        }

        private static Dictionary<string, int> ReadCells(ViewSchedule schedule, ScheduleFieldId valueId, ScheduleFieldId countId)
        {
            var visible = schedule.Definition.GetFieldOrder().Where(id => !schedule.Definition.GetField(id).IsHidden).ToList();
            int valueColumn = visible.IndexOf(valueId);
            int countColumn = visible.IndexOf(countId);
            var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            using (var body = schedule.GetTableData().GetSectionData(SectionType.Body))
            {
                if (body.NumberOfRows == 0 || body.NumberOfColumns == 0) return result;
                for (int row = body.FirstRowNumber; row <= body.LastRowNumber; row++)
                {
                    int amount;
                    string count = schedule.GetCellText(SectionType.Body, row, body.FirstColumnNumber + countColumn);
                    if (!int.TryParse(count, NumberStyles.Integer, CultureInfo.CurrentCulture, out amount) || amount <= 0) continue;
                    string text = schedule.GetCellText(SectionType.Body, row, body.FirstColumnNumber + valueColumn);
                    if (string.IsNullOrWhiteSpace(text)) continue;
                    int previous;
                    result.TryGetValue(text, out previous);
                    result[text] = previous + amount;
                }
            }
            return result;
        }

        private static void AddResult(Dictionary<string, ScheduleCopyValue> result, ScheduleCopyValue value, string display, int count)
        {
            ScheduleCopyValue existing;
            if (result.TryGetValue(value.Key, out existing)) existing.ElementCount += count;
            else
            {
                // Сохраняются только простые значения: объекты Revit не выходят за откат.
                result.Add(value.Key, new ScheduleCopyValue
                {
                    Storage = value.Storage, Text = value.Text, Integer = value.Integer, Number = value.Number,
                    ElementIdValue = value.ElementIdValue, Display = display, ElementCount = count
                });
            }
        }

        private static string CaptionKey(string text)
        {
            return (text ?? "").Replace('\u00a0', ' ').Replace('\u202f', ' ').Trim();
        }

        private static Dictionary<string, List<ScheduleCopyValue>> CollectCandidates(Document host, ScheduleDefinition definition, ScheduleField field)
        {
            var result = new Dictionary<string, List<ScheduleCopyValue>>(StringComparer.OrdinalIgnoreCase);
            var seenValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var shared = host.GetElement(field.ParameterId) as SharedParameterElement;
            var projectParameter = host.GetElement(field.ParameterId) as ParameterElement;
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
            long parameterId = field.ParameterId.IntegerValue;
#else
            long parameterId = field.ParameterId.Value;
#endif
            foreach (var document in LoadedDocuments(host))
            {
                // ID определений из разных документов не сравниваются между собой.
                Func<Element, Parameter> getter;
                if (shared != null) getter = owner => owner.get_Parameter(shared.GuidValue);
                else if (parameterId < -1) getter = owner => owner.get_Parameter((BuiltInParameter)parameterId);
                else if (document.Equals(host) && projectParameter != null)
                {
                    Definition parameterDefinition = projectParameter.GetDefinition();
                    getter = owner => owner.get_Parameter(parameterDefinition);
                }
                else continue;

                var types = new Dictionary<ElementId, Element>();
                var phases = new List<Phase> { null };
                if (field.FieldType == ScheduleFieldType.Room || field.FieldType == ScheduleFieldType.FromRoom
                    || field.FieldType == ScheduleFieldType.ToRoom || field.FieldType == ScheduleFieldType.Space)
                    phases.AddRange(document.Phases.Cast<Phase>());

                using (var collector = new FilteredElementCollector(document).WhereElementIsNotElementType())
                {
                    if (definition.CategoryId != ElementId.InvalidElementId) collector.OfCategoryId(definition.CategoryId);
                    foreach (Element element in collector)
                    foreach (Phase phase in phases)
                    foreach (Element owner in ScheduleCopyService.Owners(document, element, field.FieldType, phase))
                    {
                        var parameter = ScheduleCopyService.FindParameter(document, owner, getter,
                            field.FieldType == ScheduleFieldType.ElementType, types);
                        if (parameter == null || !parameter.HasValue) continue;
                        var value = ScheduleCopyService.ReadValue(document, parameter);
                        if (value == null) continue;
                        // Локальный ElementId связи нельзя передавать в фильтр основного файла.
                        if (value.Storage == StorageType.ElementId && !document.Equals(host)) continue;
                        if (!seenValues.Add(value.Key)) continue;

                        if (value.Storage == StorageType.Double)
                        {
                            using (var options = field.GetFormatOptions())
                            using (var format = new FormatValueOptions())
                            {
                                format.SetFormatOptions(options);
#if Debug2020 || Revit2020
                                value.Display = UnitFormatUtils.Format(host.GetUnits(), field.UnitType, value.Number, false, false, format);
#else
                                value.Display = UnitFormatUtils.Format(host.GetUnits(), field.GetSpecTypeId(), value.Number, false, format);
#endif
                            }
                        }
                        AddCandidate(result, value.Display, value);
                        if (value.Storage == StorageType.Integer)
                            AddCandidate(result, value.Integer.ToString(CultureInfo.CurrentCulture), value);
                    }
                }
            }
            return result;
        }

        private static void AddCandidate(Dictionary<string, List<ScheduleCopyValue>> result, string caption, ScheduleCopyValue value)
        {
            List<ScheduleCopyValue> items;
            string key = CaptionKey(caption);
            if (!result.TryGetValue(key, out items)) result.Add(key, items = new List<ScheduleCopyValue>());
            if (!items.Any(item => item.Key == value.Key)) items.Add(value);
        }

        private static IEnumerable<Document> LoadedDocuments(Document host)
        {
            var seen = new HashSet<Document>();
            var pending = new Queue<Document>();
            pending.Enqueue(host);
            while (pending.Count > 0)
            {
                var document = pending.Dequeue();
                if (!seen.Add(document)) continue;
                yield return document;
                using (var links = new FilteredElementCollector(document).OfClass(typeof(RevitLinkInstance)))
                {
                    foreach (RevitLinkInstance link in links)
                    {
                        var linked = link.GetLinkDocument();
                        if (linked != null && !seen.Contains(linked)) pending.Enqueue(linked);
                    }
                }
            }
        }
    }
}
