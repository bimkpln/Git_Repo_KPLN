using Autodesk.Revit.DB;
using KPLN_ViewsAndLists_Ribbon.Common.Schedules;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace KPLN_ViewsAndLists_Ribbon.Services.Schedules
{
    public static class ScheduleBatchService
    {
        public static List<BatchScheduleFilter> ReadCommonFilters(Document doc, List<ScheduleItem> sources, IEnumerable<ScheduleParameterChoice> parameters)
        {
            var result = new List<BatchScheduleFilter>();
            if (sources.Count == 0) return result;
            var used = sources.ToDictionary(item => item.Schedule.UniqueId, item => new HashSet<int>());
            foreach (var candidate in sources[0].Filters)
            {
                var field = sources[0].Schedule.Definition.GetField(candidate.Original.FieldId);
                string key = ScheduleBatchService.ParameterKey(doc, field);
                var matches = new Dictionary<string, int>();
                foreach (var source in sources)
                {
                    var match = source.Filters.FirstOrDefault(f => !used[source.Schedule.UniqueId].Contains(f.Index)
                        && ScheduleBatchService.ParameterKey(doc, source.Schedule.Definition.GetField(f.Original.FieldId)) == key
                        && candidate.Original.FilterType == f.Original.FilterType);
                    if (match == null) break;
                    matches.Add(source.Schedule.UniqueId, match.Index);
                }
                if (matches.Count != sources.Count) continue;

                var parameter = parameters.FirstOrDefault(p => p.Key == key);
                if (parameter == null)
                {
                    parameter = new ScheduleParameterChoice { Key = key, Name = field.GetName(), ScheduleCount = sources.Count, SuggestedKind = "Текст" };
                    // Существующее условие доступно для исправления, но не расширяет список новых параметров.
                }
                var original = candidate.Original;
                bool differentValues = sources.Any(source => !ScheduleService.SameValue(original,
                    source.Filters.First(f => f.Index == matches[source.Schedule.UniqueId]).Original));
                var rule = new BatchScheduleFilter
                {
                    IsExisting = true, Parameter = parameter, Condition = candidate.Condition,
                    Kind = original.IsDoubleValue ? "Число в единицах проекта" : original.IsIntegerValue ? "Целое число"
                        : original.IsElementIdValue ? "ID элемента" : "Текст",
                    Value = differentValues ? "{Исходное}" : original.IsElementIdValue ? original.GetElementIdValue().ToString() : candidate.Value
                };
                foreach (var pair in matches)
                {
                    rule.SourceIndices.Add(pair.Key, pair.Value);
                    used[pair.Key].Add(pair.Value);
                }
                rule.InitialSignature = rule.Signature;
                result.Add(rule);
            }
            return result;
        }

        public static string ParameterKey(Document doc, ScheduleField field)
        {
            var shared = doc.GetElement(field.ParameterId) as SharedParameterElement;
            return (shared == null ? field.ParameterId.ToString() : shared.GuidValue.ToString()) + "/" + field.FieldType;
        }

        internal static string ParameterKey(Document doc, SchedulableField field)
        {
            var shared = doc.GetElement(field.ParameterId) as SharedParameterElement;
            return (shared == null ? field.ParameterId.ToString() : shared.GuidValue.ToString()) + "/" + field.FieldType;
        }

        public static List<ScheduleParameterChoice> CollectParameters(Document doc, IEnumerable<ScheduleItem> schedules)
        {
            return CollectParameters(doc, schedules, new ScheduleParameterCatalog(doc));
        }

        internal static List<ScheduleParameterChoice> CollectParameters(Document doc, IEnumerable<ScheduleItem> schedules, ScheduleParameterCatalog catalog, bool includeAvailable = true)
        {
            schedules = schedules.Where(item => item.Include).GroupBy(item => item.Schedule.UniqueId).Select(group => group.First()).ToList();
            int selectedCount = schedules.Count();
            var result = new Dictionary<string, ScheduleParameterChoice>();
            foreach (ScheduleItem item in schedules)
            {
                var seen = new HashSet<string>();
                ScheduleDefinition definition = item.Schedule.Definition;
                foreach (ScheduleFieldId id in definition.GetFieldOrder())
                {
                    ScheduleField field = definition.GetField(id);
                    if (field.ParameterId == ElementId.InvalidElementId ||
                        (!definition.CanFilterByValue(id) && !definition.CanFilterBySubstring(id)
                         && !definition.CanFilterByParameterExistence(id))) continue;

                    string key = ParameterKey(doc, field);
                    if (!seen.Add(key)) continue;
                    ScheduleParameterChoice choice;
                    if (!result.TryGetValue(key, out choice))
                    {
                        choice = new ScheduleParameterChoice { Key = key, Name = field.GetName(), SuggestedKind = DetectKind(doc, field) };
                        result.Add(key, choice);
                    }
                    choice.ScheduleCount++;
                    ScheduleFilter filter = definition.GetFilters().FirstOrDefault(f => f.FieldId == id && !f.IsNullValue);
                    if (filter != null)
                        choice.SuggestedKind = filter.IsDoubleValue ? "Число в единицах проекта" : filter.IsIntegerValue ? "Целое число"
                            : filter.IsElementIdValue ? "ID элемента" : "Текст";
                }
            }
            // Полный каталог нужен только для добавления нового условия.
            if (includeAvailable)
            foreach (var item in schedules)
            {
                var definition = item.Schedule.Definition;
                var existing = new HashSet<string>(definition.GetFieldOrder().Select(id => ParameterKey(doc, definition.GetField(id))));
                foreach (var available in catalog.GetFields(definition))
                {
                    string key = ParameterKey(doc, available);
                    if (!existing.Add(key)) continue;
                    ScheduleParameterChoice choice;
                    if (!result.TryGetValue(key, out choice))
                    {
                        choice = new ScheduleParameterChoice
                        {
                            Key = key, Name = available.GetName(doc), SuggestedKind = DetectParameterKind(doc, available.ParameterId)
                        };
                        result.Add(key, choice);
                    }
                    choice.ScheduleCount++;
                }
            }
            var common = result.Values.Where(p => selectedCount > 0 && p.ScheduleCount == selectedCount)
                .OrderBy(p => p.Name).ThenBy(p => p.Key).ToList();
#if DEBUG
            if (includeAvailable) WriteParameterCatalogAudit(schedules, common);
#endif
            return common;
        }

#if DEBUG
        // Диагностика включается только локальным файлом запроса; модель не изменяется.
        private static void WriteParameterCatalogAudit(IEnumerable<ScheduleItem> schedules, IEnumerable<ScheduleParameterChoice> parameters)
        {
            string prefix = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "KPLN_ScheduleParameterCatalog_20261019");
            if (!System.IO.File.Exists(prefix + ".request")) return;

            try
            {
                using (var writer = System.Xml.XmlWriter.Create(prefix + ".xml", new System.Xml.XmlWriterSettings { Indent = true }))
                {
                    writer.WriteStartElement("catalog");
                    writer.WriteAttributeString("utc", DateTime.UtcNow.ToString("O"));
                    writer.WriteAttributeString("assembly", typeof(ScheduleBatchService).Assembly.Location);
                    writer.WriteAttributeString("module", typeof(ScheduleBatchService).Module.ModuleVersionId.ToString());
                    foreach (var item in schedules)
                    {
                        writer.WriteStartElement("schedule");
                        writer.WriteAttributeString("id", item.Schedule.UniqueId);
                        writer.WriteAttributeString("name", item.Schedule.Name);
                        writer.WriteAttributeString("category", item.Schedule.Definition.CategoryId.ToString());
                        var doc = item.Schedule.Document;
                        var definition = item.Schedule.Definition;
                        var nativeIds = new HashSet<ElementId>(TableView.GetAvailableParameters(doc, definition.CategoryId));
                        var bound = new Dictionary<ElementId, string>();
                        var iterator = doc.ParameterBindings.ForwardIterator();
                        iterator.Reset();
                        while (iterator.MoveNext())
                        {
                            var parameterDefinition = iterator.Key as InternalDefinition;
                            var binding = iterator.Current as ElementBinding;
                            if (parameterDefinition != null && binding != null)
                                bound[parameterDefinition.Id] = string.Join(",", binding.Categories.Cast<Category>().Select(category => category.Id.ToString()));
                        }
                        foreach (var field in definition.GetSchedulableFields())
                        {
                            writer.WriteStartElement("candidate");
                            writer.WriteAttributeString("id", field.ParameterId.ToString());
                            writer.WriteAttributeString("name", field.GetName(doc));
                            writer.WriteAttributeString("type", field.FieldType.ToString());
                            writer.WriteAttributeString("native", nativeIds.Contains(field.ParameterId).ToString());
                            string categories;
                            writer.WriteAttributeString("bindings", bound.TryGetValue(field.ParameterId, out categories) ? categories : "");
                            writer.WriteEndElement();
                        }
                        writer.WriteEndElement();
                    }
                    foreach (var parameter in parameters)
                    {
                        writer.WriteStartElement("parameter");
                        writer.WriteAttributeString("key", parameter.Key);
                        writer.WriteAttributeString("name", parameter.Name);
                        writer.WriteAttributeString("count", parameter.ScheduleCount.ToString());
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement();
                }
            }
            catch (System.IO.IOException ex) { System.Diagnostics.Trace.WriteLine(ex.Message); }
            catch (UnauthorizedAccessException ex) { System.Diagnostics.Trace.WriteLine(ex.Message); }
        }
#endif

        public static string FilterValuesForName(Document doc, ScheduleItem source, IEnumerable<BatchScheduleFilter> rules, ScheduleLevelRow level, int index,
            IEnumerable<ScheduleItem> selectedSources = null)
        {
            var ruleList = rules.ToList();
            var selected = (selectedSources ?? new[] { source }).ToList();
            var entries = FilterEntriesForName(doc, source, ruleList, level, index);
            if (selected.Count > 1)
            {
                // Общность определяется по параметру, условию и итоговому значению, а не по тексту значения отдельно.
                var common = new HashSet<Tuple<string, string, string>>(entries);
                foreach (var other in selected.Where(s => s.Schedule.UniqueId != source.Schedule.UniqueId))
                    common.IntersectWith(FilterEntriesForName(doc, other, ruleList, level, index));
                entries.RemoveAll(entry => common.Contains(entry));
            }
            return string.Join("/", entries.Select(entry => entry.Item3).Distinct(StringComparer.Ordinal));
        }

        internal static List<Tuple<string, string, string>> FilterEntriesForName(Document doc, ScheduleItem source, IEnumerable<BatchScheduleFilter> rules, ScheduleLevelRow level, int index, ScheduleOperation operation = null)
        {
            var values = new List<string>();
            var keys = new List<string>();
            var conditions = new List<string>();
            var definition = source.Schedule.Definition;
            foreach (var filter in definition.GetFilters())
            {
                var field = definition.GetField(filter.FieldId);
                keys.Add(ParameterKey(doc, field));
                conditions.Add(filter.FilterType.ToString());
                values.Add(filter.IsNullValue ? "" : filter.IsElementIdValue ? filter.GetElementIdValue().ToString()
                    : ScheduleService.DisplayValue(doc, field, filter));
            }
            foreach (var rule in rules.Where(r => !r.IsAutomatic && r.HasChanges))
            {
                int originalPosition;
                // Исходные позиции сохраняются: удаляем значение, а не строку списка.
                if (rule.Deleted)
                {
                    if (rule.SourceIndices.TryGetValue(source.Schedule.UniqueId, out originalPosition)) values[originalPosition] = "";
                    continue;
                }
                if (rule.Condition == "Параметр существует" || rule.Condition == "Есть значение" || rule.Condition == "Нет значения")
                {
                    if (rule.SourceIndices.TryGetValue(source.Schedule.UniqueId, out originalPosition)) values[originalPosition] = "";
                    continue;
                }
                if (!rule.IsExisting && (rule.Value ?? "").Contains("{Исходное}"))
                    throw new InvalidOperationException("Маска {Исходное} доступна только для существующих фильтров. Для нового фильтра задайте конкретное значение.");
                if (rule.Parameter == null) throw new InvalidOperationException("Выберите параметр для нового правила.");
                ScheduleField field;
                ScheduleFilter original = null;
                if (rule.IsExisting)
                {
                    int position;
                    if (!rule.SourceIndices.TryGetValue(source.Schedule.UniqueId, out position)) continue;
                    original = definition.GetFilter(position);
                    field = definition.GetField(original.FieldId);
                }
                else
                {
                    field = definition.GetFieldOrder().Select(definition.GetField)
                        .FirstOrDefault(f => ParameterKey(doc, f) == rule.Parameter.Key);
                    if (field == null)
                    {
                        if (!definition.GetSchedulableFields().Any(f => ParameterKey(doc, f) == rule.Parameter.Key)) continue;
                        if ((rule.Value ?? "").Contains("{Исходное}"))
                            throw new InvalidOperationException("Для нового поля нет исходного фильтра: " + rule.Parameter.Name);
                        values.Add(ExpandRule(rule.Value, level, "", index, operation));
                        keys.Add(rule.Parameter.Key);
                        conditions.Add(GetCondition(rule.Condition).ToString());
                        continue;
                    }
                    var matches = definition.GetFilters().Where(f => f.FieldId == field.FieldId).ToList();
                    if (matches.Count == 1) original = matches[0];
                }
                if (original == null && (rule.Value ?? "").Contains("{Исходное}"))
                    throw new InvalidOperationException("Для {Исходное} нужен ровно один исходный фильтр: " + field.GetName());
                string oldValue = original == null ? "" : original.IsElementIdValue ? original.GetElementIdValue().ToString()
                    : ScheduleService.DisplayValue(doc, field, original);
                string value = ExpandRule(rule.Value, level, oldValue, index, operation);
                if (rule.IsExisting && rule.SourceIndices.TryGetValue(source.Schedule.UniqueId, out originalPosition))
                {
                    values[originalPosition] = value;
                    conditions[originalPosition] = GetCondition(rule.Condition).ToString();
                }
                else
                {
                    values.Add(value);
                    keys.Add(ParameterKey(doc, field));
                    conditions.Add(GetCondition(rule.Condition).ToString());
                }
            }
            var entries = values.Select((v, i) => Tuple.Create(keys[i], conditions[i], v))
                .Where(entry => !string.IsNullOrWhiteSpace(entry.Item3)).ToList();
            if (operation?.CopyValue != null)
            {
                entries.RemoveAll(entry => entry.Item1 == operation.CopyParameter.Key);
                entries.Add(Tuple.Create(operation.CopyParameter.Key, "Equal", operation.CopyValue.Display));
            }
            return entries;
        }

        private static bool UsesOriginalValue(BatchScheduleFilter rule)
        {
            var levelRule = rule as LevelScheduleFilter;
            return (levelRule == null || (levelRule.NeedsValue && levelRule.Mode == FilterModes.Pattern))
                && (rule.Value ?? "").Contains("{Исходное}");
        }

        private static string ResolveRuleValue(Document doc, BatchScheduleFilter rule, ScheduleLevelRow level,
            string original, int index, ScheduleOperation operation)
        {
            var levelRule = rule as LevelScheduleFilter;
            return levelRule == null ? ExpandRule(rule.Value, level, original, index, operation)
                : ScheduleLevelFilterService.ResolveValue(doc, levelRule, level, index);
        }

        private static string ExpandRule(string pattern, ScheduleLevelRow level, string original, int index, ScheduleOperation operation)
        {
            return ScheduleNaming.Expand(pattern, level.Floor, level.Name, level.Elevation, level.Group, original, index,
                null, operation?.CopyParameter?.Name, operation?.CopyValue?.Display);
        }

        private static string StorageKind(StorageType storage)
        {
            switch (storage)
            {
                case StorageType.String: return "Текст";
                case StorageType.Integer: return "Целое число";
                case StorageType.Double: return "Число в единицах проекта";
                case StorageType.ElementId: return "ID элемента";
                default: return "Не определён";
            }
        }

        private static string DetectKind(Document doc, ScheduleField field)
        {
#if !Debug2020 && !Revit2020
            string kind = SpecKind(field.GetSpecTypeId());
            if (kind != "Не определён") return kind;
#endif
            return DetectParameterKind(doc, field.ParameterId);
        }

#if !Debug2020 && !Revit2020
        private static string SpecKind(ForgeTypeId spec)
        {
            if (spec == SpecTypeId.String.Text || spec == SpecTypeId.String.MultilineText || spec == SpecTypeId.String.Url) return "Текст";
            if (spec == SpecTypeId.Int.Integer || spec == SpecTypeId.Int.NumberOfPoles || spec == SpecTypeId.Boolean.YesNo) return "Целое число";
            if (UnitUtils.IsMeasurableSpec(spec)) return "Число в единицах проекта";
            if (spec == SpecTypeId.Reference.Material || spec == SpecTypeId.Reference.Image
                || spec == SpecTypeId.Reference.LoadClassification || spec == SpecTypeId.Reference.FillPattern
                || spec.TypeId.StartsWith("autodesk.revit.category:", StringComparison.Ordinal)) return "ID элемента";
            return "Не определён";
        }
#endif

        private static string DetectParameterKind(Document doc, ElementId id)
        {
            var parameter = doc.GetElement(id) as ParameterElement;
#if Debug2020 || Revit2020
            if (parameter == null) return "Не определён";
            switch (parameter.GetDefinition().ParameterType)
            {
                case ParameterType.Text:
                case ParameterType.MultilineText:
                case ParameterType.URL: return "Текст";
                case ParameterType.Integer:
                case ParameterType.YesNo: return "Целое число";
                case ParameterType.Material:
                case ParameterType.FamilyType:
                case ParameterType.Image: return "ID элемента";
                case ParameterType.Invalid: return "Не определён";
                default: return "Число в единицах проекта";
            }
#else
            if (parameter != null) return SpecKind(parameter.GetDefinition().GetDataType());
#if Debug2023 || Revit2023
            long value = id.IntegerValue;
#else
            long value = id.Value;
#endif
            if (value < -1 && Enum.IsDefined(typeof(BuiltInParameter), (int)value))
                return StorageKind(doc.GetTypeOfStorage(ParameterUtils.GetParameterTypeId((BuiltInParameter)value)));
            return "Не определён";
#endif
        }

        private static ScheduleFilterType GetCondition(string name)
        {
            string[] names = { "Equal", "NotEqual", "GreaterThan", "GreaterThanOrEqual", "LessThan", "LessThanOrEqual",
                "Contains", "NotContains", "BeginsWith", "NotBeginsWith", "EndsWith", "NotEndsWith", "HasParameter", "HasValue", "HasNoValue", "IsAssociatedWithGlobalParameter", "IsNotAssociatedWithGlobalParameter" };
            int index = Array.IndexOf(BatchScheduleFilter.Conditions, name);
            if (index < 0 || index >= names.Length) throw new InvalidOperationException("Неизвестное условие: " + name);
            return (ScheduleFilterType)Enum.Parse(typeof(ScheduleFilterType), names[index]);
        }

        private static ScheduleFilter MakeFilter(Document doc, ScheduleField field, BatchScheduleFilter rule, string text)
        {
            ScheduleFilterType condition = GetCondition(rule.Condition);
            if (condition == ScheduleFilterType.HasParameter || condition.ToString() == "HasValue" || condition.ToString() == "HasNoValue") return new ScheduleFilter(field.FieldId, condition);
            if (rule.Kind == "Не определён")
                throw new InvalidOperationException("Не удалось определить тип параметра «" + field.GetName() + "». Изменение фильтра не выполнено.");
            if (rule.Kind == "Текст") return new ScheduleFilter(field.FieldId, condition, text);
            int integer;
            if (rule.Kind == "Целое число" && int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out integer))
                return new ScheduleFilter(field.FieldId, condition, integer);
            if (rule.Kind == "Число в единицах проекта")
            {
                double number;
#if Debug2020 || Revit2020
                bool valid = UnitFormatUtils.TryParse(doc.GetUnits(), field.UnitType, text, out number);
#else
                bool valid = UnitFormatUtils.TryParse(doc.GetUnits(), field.GetSpecTypeId(), text, out number);
#endif
                if (valid) return new ScheduleFilter(field.FieldId, condition, number);
            }
            if (rule.Kind == "ID элемента")
            {
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                int id;
                if (int.TryParse(text, out id) && doc.GetElement(new ElementId(id)) != null)
#else
                long id;
                if (long.TryParse(text, out id) && doc.GetElement(new ElementId(id)) != null)
#endif
                    return new ScheduleFilter(field.FieldId, condition, new ElementId(id));
            }
            throw new InvalidOperationException("Значение «" + text + "» не соответствует типу «" + rule.Kind + "»: " + field.GetName());
        }

        public static void ApplyBatchFilters(Document doc, ViewSchedule schedule, IEnumerable<BatchScheduleFilter> rules,
            ScheduleLevelRow level, int index, List<string> messages, string sourceKey = "")
        {
            ApplyBatchFilters(doc, schedule, rules, level, index, messages, sourceKey, null);
        }

        private static void ApplyBatchFilters(Document doc, ViewSchedule schedule, IEnumerable<BatchScheduleFilter> rules,
            ScheduleLevelRow level, int index, List<string> messages, string sourceKey, ScheduleOperation operation)
        {
            ScheduleDefinition definition = schedule.Definition;
            var filters = definition.GetFilters().ToList();
            bool changed = false;
            // Индексы относятся к исходной спецификации; удаляем с конца, чтобы они не смещались.
            foreach (BatchScheduleFilter rule in rules.Where(r => !r.IsAutomatic && r.IsExisting && r.HasChanges)
                .OrderByDescending(r => r.SourceIndices.ContainsKey(sourceKey) ? r.SourceIndices[sourceKey] : -1))
            {
                int position;
                if (!rule.SourceIndices.TryGetValue(sourceKey, out position) || position >= filters.Count)
                    throw new InvalidOperationException("Состав фильтров изменился. Повторно выберите спецификации.");

                ScheduleFilter original = filters[position];
                ScheduleField field = definition.GetField(original.FieldId);
                if (rule.Deleted)
                {
                    filters.RemoveAt(position);
                    messages.Add(field.GetName() + ": удалён общий фильтр «" + rule.Condition + "».");
                }
                else
                {
                    string oldValue = original.IsElementIdValue ? original.GetElementIdValue().ToString()
                        : ScheduleService.DisplayValue(doc, field, original);
                    string text = ResolveRuleValue(doc, rule, level, oldValue, index, operation);
                    filters[position] = MakeFilter(doc, field, rule, text);
                    messages.Add(field.GetName() + ": " + rule.Condition + " «" + text + "».");
                }
                changed = true;
            }

            // Новые строки добавляют условия, не удаляя остальные фильтры параметра.
            var newRules = rules.Where(r => !r.IsAutomatic && !r.IsExisting && r.Enabled).ToList();
            var availableFields = newRules.Count == 0 ? null : definition.GetSchedulableFields()
                .GroupBy(field => ParameterKey(doc, field)).ToDictionary(group => group.Key, group => group.First());
            foreach (BatchScheduleFilter rule in newRules)
            {
                if (!rule.IsExisting && UsesOriginalValue(rule))
                    throw new InvalidOperationException("Маска {Исходное} доступна только для существующих фильтров. Для нового фильтра задайте конкретное значение.");
                if (rule.Parameter == null) throw new InvalidOperationException("Выберите параметр для нового правила.");
                SchedulableField available;
                if (!availableFields.TryGetValue(rule.Parameter.Key, out available))
                    throw new InvalidOperationException("Параметр «" + rule.Parameter.Name + "» недоступен для спецификации «" + schedule.Name + "». Повторно выберите параметр.");

                var fields = definition.GetFieldOrder().Select(definition.GetField)
                    .Where(f => ParameterKey(doc, f) == rule.Parameter.Key).ToList();
                if (fields.Count == 0)
                {
                    if (available != null)
                    {
                        ScheduleService.EnsureEditableFilters(doc, schedule, level.Name, messages);
                        var added = definition.AddField(available);
                        added.IsHidden = true;
                        fields.Add(added);
                        messages.Add("Добавлено скрытое поле «" + added.GetName() + "».");
                    }
                }
                if (fields.Count == 0)
                {
                    messages.Add("Нет параметра «" + rule.Parameter.Name + "»; новое правило пропущено.");
                    continue;
                }
                foreach (ScheduleField field in fields)
                {
                    var oldFilters = definition.GetFilters().Where(f => f.FieldId == field.FieldId).ToList();
                    if (oldFilters.Count != 1 && UsesOriginalValue(rule))
                        throw new InvalidOperationException("Для {Исходное} нужен ровно один исходный фильтр: " + field.GetName());
                    string original = oldFilters.Count == 1 ? (oldFilters[0].IsElementIdValue
                        ? oldFilters[0].GetElementIdValue().ToString() : ScheduleService.DisplayValue(doc, field, oldFilters[0])) : "";
                    string text = ResolveRuleValue(doc, rule, level, original, index, operation);
                    filters.Add(MakeFilter(doc, field, rule, text));
                    messages.Add(field.GetName() + ": добавлен фильтр " + rule.Condition + " «" + text + "».");
                    changed = true;
                }
            }
            if (operation?.CopyValue != null)
            {
                // Заменяем условия копирования до единственного SetFilters: промежуточный
                // набор со старыми условиями мог бы превысить ограничение Revit.
                filters = ScheduleCopyService.WithCopyFilter(doc, schedule, operation, filters, messages);
                changed = true;
            }
            if (!changed) return;

            ScheduleService.EnsureEditableFilters(doc, schedule, level.Name, messages);
            definition.SetFilters(filters);
            doc.Regenerate();
            var actual = definition.GetFilters();
            if (actual.Count != filters.Count || actual.Where((f, i) => f.FieldId != filters[i].FieldId || !ScheduleService.SameValue(f, filters[i])).Any())
                throw new InvalidOperationException("Revit не сохранил заданные фильтры: " + schedule.Name);
            messages.Add("Фильтров после замены: " + filters.Count + ".");
        }

        // Предпросмотр выполняет те же операции внутри транзакции с обязательным откатом.
        public static bool Preview(Document doc, IList<ScheduleOperation> operations, IList<BatchScheduleFilter> rules)
        {
            ValidateCopySources(operations);
            bool valid = true;
            using (var transaction = new Transaction(doc, "KPLN: проверка спецификаций"))
            {
                transaction.Start();
                try
                {
                    foreach (ScheduleOperation operation in operations)
                    {
                        using (var sub = new SubTransaction(doc))
                        {
                            sub.Start();
                            try
                            {
                                var messages = new List<string>();
                                ExecuteOne(doc, operation, rules, messages);
                                if (sub.Commit() != TransactionStatus.Committed)
                                    throw new InvalidOperationException("Revit откатил проверяемую операцию.");
                                operation.Result = "Готово. " + string.Join(" ", messages);
                            }
                            catch (Exception ex)
                            {
                                if (sub.GetStatus() == TransactionStatus.Started) sub.RollBack();
                                operation.Result = "Ошибка: " + ex.Message;
                                valid = false;
                            }
                        }
                    }
                }
                finally { transaction.RollBack(); }
            }
            return valid;
        }

        public static int Execute(Document doc, IList<ScheduleOperation> operations, IList<BatchScheduleFilter> rules)
        {
            foreach (var operation in operations)
            {
                operation.Committed = null;
                operation.WarningCount = 0;
                operation.Result = "Не выполнено.";
            }
            ValidateCopySources(operations);
            int done = 0;
            try
            {
                using (var group = new TransactionGroup(doc, "KPLN: Конструктор спецификаций"))
                {
                    group.Start();
                    foreach (ScheduleOperation operation in operations)
                    {
                        using (var transaction = new Transaction(doc, "Спецификация: " + operation.TargetName))
                        {
                            transaction.Start();
                            var messages = new List<string>();
                            transaction.SetFailureHandlingOptions(transaction.GetFailureHandlingOptions()
                                .SetFailuresPreprocessor(new ScheduleFailures(messages, operation)).SetClearAfterRollback(true));
                            try
                            {
                                ExecuteOne(doc, operation, rules, messages);
                                if (transaction.Commit() != TransactionStatus.Committed)
                                    throw new InvalidOperationException("Revit откатил изменения. " + string.Join(" ", messages));
                                done++;
                                operation.Committed = true;
                                operation.Result = "Выполнено. " + string.Join(" ", messages);
                            }
                            catch (Exception ex)
                            {
                                if (transaction.GetStatus() == TransactionStatus.Started) transaction.RollBack();
                                operation.Committed = false;
                                operation.Result = "Откат: " + ex.Message;
                            }
                        }
                    }
                    if (done > 0)
                    {
                        if (group.Assimilate() != TransactionStatus.Committed)
                            throw new InvalidOperationException("Revit не сохранил пакет изменений.");
                    }
                    else group.RollBack();
                }
                }
            catch
            {
                // При откате всей группы промежуточные Commit не считаются результатом.
                foreach (var operation in operations.Where(item => item.Committed == true))
                {
                    operation.Committed = false;
                    operation.Result = "Откат всего пакета. Изменения не сохранены.";
                }
                throw;
            }
            return done;
        }

        private static void ExecuteOne(Document doc, ScheduleOperation operation, IList<BatchScheduleFilter> rules, List<string> messages)
        {
            ViewSchedule target;
            if (operation.CreateCopy)
                target = ScheduleService.Create(doc, operation.Source, operation.Level, operation.Index, messages, operation.TargetName);
            else
            {
                target = operation.Source.Schedule;
                ScheduleService.ApplyFilters(doc, target, target, operation.Source.Filters, operation.Level, operation.Index, messages);
            }
            ApplyBatchFilters(doc, target, rules, operation.Level, operation.Index, messages, operation.Source.Schedule.UniqueId, operation);
        }

        private static void ValidateCopySources(IList<ScheduleOperation> operations)
        {
            if (operations.Any(operation => operation.CreateCopy) &&
                operations.Select(operation => operation.Source.Schedule.UniqueId).Distinct().Count() != 1)
                throw new InvalidOperationException("Для создания копий должна быть выбрана ровно одна исходная спецификация.");
        }

        private sealed class ScheduleFailures : IFailuresPreprocessor
        {
            private readonly List<string> _messages;
            private readonly ScheduleOperation _operation;
            public ScheduleFailures(List<string> messages, ScheduleOperation operation)
            {
                _messages = messages;
                _operation = operation;
            }
            public FailureProcessingResult PreprocessFailures(FailuresAccessor accessor)
            {
                bool errors = false;
                foreach (FailureMessageAccessor message in accessor.GetFailureMessages())
                {
                    _messages.Add(message.GetDescriptionText());
                    if (message.GetSeverity() == FailureSeverity.Warning)
                    {
                        _operation.WarningCount++;
                        accessor.DeleteWarning(message);
                    }
                    else errors = true;
                }
                return errors ? FailureProcessingResult.ProceedWithRollBack : FailureProcessingResult.Continue;
            }
        }
    }
}
