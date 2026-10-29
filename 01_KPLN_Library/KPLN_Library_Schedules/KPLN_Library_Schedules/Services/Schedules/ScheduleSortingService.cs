using Autodesk.Revit.DB;
using KPLN_Library_Schedules.Common.Schedules;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace KPLN_Library_Schedules.Services.Schedules
{
    internal static class ScheduleSortingService
    {
        internal static BatchScheduleSort Clone(BatchScheduleSort source)
        {
            var result = new BatchScheduleSort
            {
                IsExisting = source.IsExisting, Enabled = source.Enabled, Deleted = source.Deleted,
                Parameter = source.Parameter, Position = source.Position, Direction = source.Direction,
                Visibility = source.Visibility, InitialParameterKey = source.InitialParameterKey,
                InitialPosition = source.InitialPosition, InitialVisibility = source.InitialVisibility,
                InitialSignature = source.InitialSignature
            };
            foreach (var pair in source.SourceIndices) result.SourceIndices.Add(pair.Key, pair.Value);
            return result;
        }

        internal static XElement Save(ScheduleItem source)
        {
            return source.LevelSorts == null ? null : SaveRules(source.LevelSorts, source.Schedule.UniqueId);
        }

        internal static XElement SaveRules(IEnumerable<BatchScheduleSort> rules, string sourceKey)
        {
            return new XElement("Sorting", rules.Select(rule => new XElement("Rule",
                new XAttribute("Existing", rule.IsExisting), new XAttribute("Enabled", rule.Enabled),
                new XAttribute("Deleted", rule.Deleted), new XAttribute("Position", rule.Position ?? "Исходное"),
                new XAttribute("Direction", rule.Direction), new XAttribute("Visibility", rule.Visibility),
                new XAttribute("ParameterKey", rule.Parameter?.Key ?? ""), new XAttribute("ParameterName", rule.Parameter?.Name ?? ""),
                new XAttribute("InitialKey", rule.InitialParameterKey ?? ""), new XAttribute("InitialSignature", rule.InitialSignature ?? ""),
                new XAttribute("Index", rule.SourceIndices.ContainsKey(sourceKey) ? rule.SourceIndices[sourceKey] : -1))));
        }

        internal static List<BatchScheduleSort> Load(Document doc, ScheduleItem source, XElement saved)
        {
            return saved == null ? null : LoadRules(ReadCommon(doc, new List<ScheduleItem> { source }), source.Schedule.UniqueId, saved);
        }

        internal static List<BatchScheduleSort> LoadRules(List<BatchScheduleSort> current, string sourceKey, XElement saved)
        {
            var result = current.Select(Clone).ToList();
            foreach (var item in saved.Elements("Rule"))
            {
                bool existing = (bool?)item.Attribute("Existing") ?? false;
                var row = existing ? result.FirstOrDefault(rule => rule.InitialParameterKey == (string)item.Attribute("InitialKey")
                    && rule.SourceIndices.ContainsKey(sourceKey) && rule.SourceIndices[sourceKey] == (int?)item.Attribute("Index")
                    && rule.InitialSignature == (string)item.Attribute("InitialSignature")) : new BatchScheduleSort();
                if (row == null) continue;
                row.Parameter = new ScheduleParameterChoice { Key = (string)item.Attribute("ParameterKey"), Name = (string)item.Attribute("ParameterName"), ScheduleCount = 1 };
                row.Position = (string)item.Attribute("Position") ?? "Исходное";
                row.Direction = (string)item.Attribute("Direction") ?? "По возрастанию";
                row.Visibility = (string)item.Attribute("Visibility") ?? "Исходное";
                row.Enabled = (bool?)item.Attribute("Enabled") ?? true;
                row.Deleted = (bool?)item.Attribute("Deleted") ?? false;
                if (!existing) result.Add(row);
            }
            return result;
        }

        internal static string Key(Document doc, ScheduleField field)
        {
            return field.ParameterId == ElementId.InvalidElementId
                ? "field:" + field.GetName() + "/" + field.FieldType
                : ScheduleBatchService.ParameterKey(doc, field);
        }

        internal static string ReadVisibility(IEnumerable<ScheduleField> fields)
        {
            var values = fields.Select(field => field.IsHidden).Distinct().ToList();
            return values.Count != 1 ? "Исходное" : values[0] ? "Скрыт" : "Виден";
        }

        internal static List<BatchScheduleSort> ReadCommon(Document doc, List<ScheduleItem> sources)
        {
            var result = new List<BatchScheduleSort>();
            if (sources.Count == 0) return result;

            var sourceSorts = sources.Select(source => source.Schedule.Definition.GetSortGroupFields()
                .Select((sort, index) => new { Sort = sort, Index = index, Field = source.Schedule.Definition.GetField(sort.FieldId) })
                .ToList()).ToList();
            var keys = sourceSorts.Select(items => (IList<string>)items.Select(item => Key(doc, item.Field)).ToList()).ToList();

            // Повторения сопоставляем по порядку: каждое вхождение используется только один раз.
            foreach (var positions in ScheduleSortOrderPlanner.MatchCommon(keys))
            {
                var matches = sourceSorts.Select((items, i) => items[positions[i]]).ToList();
                var original = matches[0].Sort;
                var field = matches[0].Field;
                string key = Key(doc, field);

                var row = new BatchScheduleSort
                {
                    IsExisting = true,
                    Parameter = new ScheduleParameterChoice { Key = key, Name = field.GetName(), ScheduleCount = sources.Count },
                    Position = matches.Select(item => item.Index).Distinct().Count() == 1
                        ? (matches[0].Index + 1).ToString() : "Исходное",
                    Direction = matches.Select(item => item.Sort.SortOrder).Distinct().Count() == 1
                        ? original.SortOrder == ScheduleSortOrder.Ascending ? "По возрастанию" : "По убыванию" : "Исходное",
                    Visibility = ReadVisibility(matches.Select(item => item.Field))
                };
                for (int i = 0; i < sources.Count; i++) row.SourceIndices.Add(sources[i].Schedule.UniqueId, matches[i].Index);
                row.InitialPosition = row.Position;
                row.InitialParameterKey = row.Parameter.Key;
                row.InitialVisibility = row.Visibility;
                row.InitialSignature = row.Signature;
                result.Add(row);
            }
            return result;
        }

        // Проверяем доступность сортировки штатным API; временные поля и шаблоны откатываются.
        internal static List<ScheduleParameterChoice> CollectParameters(Document doc, List<ScheduleItem> sources, ScheduleParameterCatalog catalog)
        {
            var common = new Dictionary<string, ScheduleParameterChoice>();
            if (sources.Count == 0) return common.Values.ToList();
            var candidates = sources.ToDictionary(source => source, source => catalog.GetFields(source.Schedule.Definition));
            using (var transaction = new Transaction(doc, "KPLN: чтение параметров сортировки"))
            {
                transaction.Start();
                try
                {
                    foreach (var source in sources)
                    {
                        using (var sub = new SubTransaction(doc))
                        {
                            sub.Start();
                            try
                            {
                                EnsureEditable(doc, source.Schedule, new List<string>());
                                var definition = source.Schedule.Definition;
                                var fields = definition.GetFieldOrder().Select(definition.GetField).ToList();
                                var known = new HashSet<string>(fields.Select(field => Key(doc, field)));
                                foreach (var candidate in candidates[source])
                                    if (known.Add(ScheduleBatchService.ParameterKey(doc, candidate))) fields.Add(definition.AddField(candidate));

                                foreach (var group in fields.Where(field => definition.CanSortByField(field.FieldId)).GroupBy(field => Key(doc, field)))
                                {
                                    ScheduleParameterChoice choice;
                                    if (!common.TryGetValue(group.Key, out choice))
                                    {
                                        choice = new ScheduleParameterChoice { Key = group.Key, Name = group.First().GetName() };
                                        common.Add(group.Key, choice);
                                    }
                                    choice.ScheduleCount++;
                                }
                            }
                            finally { sub.RollBack(); }
                        }
                    }
                }
                finally { transaction.RollBack(); }
            }
            return common.Values.Where(choice => choice.ScheduleCount == sources.Count).OrderBy(choice => choice.Name).ToList();
        }

        internal static void EnsureEditable(Document doc, ViewSchedule schedule, List<string> messages)
        {
            var parameters = new[] { BuiltInParameter.SCHEDULE_FIELDS_PARAM, BuiltInParameter.SCHEDULE_GROUP_PARAM, BuiltInParameter.SCHEDULE_FILTER_PARAM };
            if (!parameters.Any(parameter => ScheduleSupport.TemplateControls(doc, schedule, parameter))) return;

            var template = (View)doc.GetElement(schedule.ViewTemplateId);
            var own = (View)doc.GetElement(template.Duplicate(ViewDuplicateOption.Duplicate));
            own.Name = ScheduleSupport.UniqueViewName(doc, template.Name + " — " + schedule.Name);
            var free = own.GetNonControlledTemplateParameterIds().ToList();
            foreach (var parameter in parameters)
            {
                var id = new ElementId(parameter);
                if (!free.Contains(id)) free.Add(id);
            }
            own.SetNonControlledTemplateParameterIds(free);
            schedule.ViewTemplateId = own.Id;
            messages.Add("Создан индивидуальный шаблон для полей и сортировки: " + own.Name);
        }

        internal static void Apply(Document doc, ViewSchedule schedule, string sourceKey, IList<BatchScheduleSort> rules, List<string> messages)
        {
            if (rules == null || !rules.Any(rule => rule.HasChanges)) return;
            EnsureEditable(doc, schedule, messages);
            var definition = schedule.Definition;
            var sorts = definition.GetSortGroupFields().ToList();
            var originalSorts = sorts.ToList();
            var moved = new Dictionary<int, ScheduleSortGroupField>();
            var newKeys = new HashSet<string>();

            foreach (var rule in rules.Where(row => row.HasChanges)
                .OrderBy(row => row.IsExisting && row.Deleted ? 0 : row.IsExisting ? 1 : 2))
            {
                if (rule.Parameter == null) throw new InvalidOperationException("Выберите параметр сортировки.");
                ScheduleSortGroupField sort;
                if (rule.IsExisting)
                {
                    int position;
                    if (!rule.SourceIndices.TryGetValue(sourceKey, out position) || position < 0 || position >= originalSorts.Count
                        || Key(doc, definition.GetField(originalSorts[position].FieldId)) != rule.InitialParameterKey)
                        throw new InvalidOperationException("Состав сортировки изменился. Повторно выберите спецификации.");
                    // Один FieldId может встречаться несколько раз; выбираем конкретную исходную строку.
                    sort = originalSorts[position];
                    if (rule.Deleted) { sorts.Remove(sort); continue; }
                    if (rule.ParameterChanged)
                    {
                        string previousName = definition.GetField(sort.FieldId).GetName();
                        // Меняем только поле: заголовки, нижние колонтитулы и пустые строки сохраняются.
                        sort.FieldId = ResolveField(doc, definition, rule.Parameter).FieldId;
                        messages.Add("Сортировка: «" + previousName + "» заменён на «" + rule.Parameter.Name + "».");
                    }
                }
                else
                {
                    if (!newKeys.Add(rule.Parameter.Key) || sorts.Any(item => Key(doc, definition.GetField(item.FieldId)) == rule.Parameter.Key))
                        throw new InvalidOperationException("Параметр сортировки уже добавлен: " + rule.Parameter.Name);
                    var field = ResolveField(doc, definition, rule.Parameter);
                    sort = new ScheduleSortGroupField(field.FieldId);
                    sorts.Add(sort);
                }

                if (rule.Direction != "Исходное")
                    sort.SortOrder = rule.Direction == "По убыванию" ? ScheduleSortOrder.Descending : ScheduleSortOrder.Ascending;
                if (rule.Position != "Исходное" && (!rule.IsExisting || rule.Position != rule.InitialPosition))
                {
                    int rank;
                    if (!int.TryParse(rule.Position, out rank) || rank < 1) throw new InvalidOperationException("Индекс сортировки должен быть целым числом от 1.");
                    if (moved.ContainsKey(rank)) throw new InvalidOperationException("Повторяется индекс сортировки: " + rank);
                    moved.Add(rank, sort);
                }
            }

            // Явно заданные позиции занимают свои места; остальные сохраняют взаимный порядок.
            var ordered = ScheduleSortOrderPlanner.Arrange(sorts, moved);
            definition.SetSortGroupFields(ordered);
            doc.Regenerate();
            var actual = definition.GetSortGroupFields();
            if (actual.Count != ordered.Count || actual.Where((sort, i) => sort.FieldId != ordered[i].FieldId || sort.SortOrder != ordered[i].SortOrder).Any())
                throw new InvalidOperationException("Revit не сохранил порядок сортировки.");
            messages.Add("Сортировка: " + string.Join(" → ", ordered.Select((sort, i) => (i + 1) + ". " + definition.GetField(sort.FieldId).GetName())) + ".");
        }

        private static ScheduleField ResolveField(Document doc, ScheduleDefinition definition, ScheduleParameterChoice parameter)
        {
            var field = definition.GetFieldOrder().Select(definition.GetField)
                .FirstOrDefault(item => Key(doc, item) == parameter.Key && definition.CanSortByField(item.FieldId));
            if (field == null)
            {
                var available = definition.GetSchedulableFields().FirstOrDefault(item => ScheduleBatchService.ParameterKey(doc, item) == parameter.Key)
                    ?? throw new InvalidOperationException("Параметр сортировки недоступен: " + parameter.Name);
                field = definition.AddField(available);
                field.IsHidden = true;
            }
            if (!definition.CanSortByField(field.FieldId))
                throw new InvalidOperationException("По этому параметру нельзя сортировать: " + parameter.Name);
            return field;
        }

        internal static void ApplyVisibility(Document doc, ViewSchedule schedule, IEnumerable<BatchScheduleFilter> filters, IEnumerable<BatchScheduleSort> sorts, List<string> messages)
        {
            var requests = filters.Where(rule => rule.VisibilityChanged && rule.Parameter != null)
                .Select(rule => new { rule.Parameter.Key, rule.Visibility })
                .Concat((sorts ?? Enumerable.Empty<BatchScheduleSort>()).Where(rule => rule.VisibilityChanged && rule.Parameter != null)
                    .Select(rule => new { rule.Parameter.Key, rule.Visibility })).GroupBy(item => item.Key).ToList();
            if (requests.Count == 0) return;
            if (requests.Any(group => group.Select(item => item.Visibility).Distinct().Count() != 1))
                throw new InvalidOperationException("Для одного параметра задана разная видимость. Согласуйте строки фильтров и сортировки.");

            EnsureEditable(doc, schedule, messages);
            foreach (var request in requests)
            {
                bool hidden = request.First().Visibility == "Скрыт";
                var fields = schedule.Definition.GetFieldOrder().Select(schedule.Definition.GetField).Where(field => Key(doc, field) == request.Key).ToList();
                if (fields.Count == 0) throw new InvalidOperationException("Не найден столбец для изменения видимости: " + request.Key);
                foreach (var field in fields) field.IsHidden = hidden;
                doc.Regenerate();
                if (fields.Any(field => field.IsHidden != hidden)) throw new InvalidOperationException("Revit не сохранил видимость столбца.");
                messages.Add(fields[0].GetName() + (hidden ? ": столбец скрыт." : ": столбец виден."));
            }
        }
    }

    internal static class ScheduleSortOrderPlanner
    {
        internal static List<int[]> MatchCommon(IList<IList<string>> sources)
        {
            var result = new List<int[]>();
            if (sources.Count == 0) return result;
            var occurrences = new Dictionary<string, int>();
            for (int firstIndex = 0; firstIndex < sources[0].Count; firstIndex++)
            {
                string key = sources[0][firstIndex];
                int occurrence;
                occurrences.TryGetValue(key, out occurrence);
                occurrences[key] = occurrence + 1;
                var positions = sources.Select(items => Enumerable.Range(0, items.Count)
                    .Where(index => items[index] == key).Skip(occurrence).DefaultIfEmpty(-1).First()).ToArray();
                if (positions.All(index => index >= 0)) result.Add(positions);
            }
            return result;
        }

        internal static List<T> Arrange<T>(IList<T> items, IDictionary<int, T> moved)
        {
            if (moved.Keys.Any(rank => rank < 1 || rank > items.Count))
                throw new InvalidOperationException("Индекс сортировки должен быть от 1 до " + items.Count + ".");
            var untouched = new Queue<T>(items.Where(item => !moved.Values.Contains(item)));
            return Enumerable.Range(1, items.Count).Select(rank => moved.ContainsKey(rank) ? moved[rank] : untouched.Dequeue()).ToList();
        }
    }
}
