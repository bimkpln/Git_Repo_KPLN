using Autodesk.Revit.DB;
using KPLN_Library_Schedules.Common.Schedules;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace KPLN_Library_Schedules.Services.Schedules
{
    public static class ScheduleLevelFilterService
    {
        public static List<LevelScheduleFilter> ReadRules(Document doc, ScheduleItem source)
        {
            if (source.LevelFilters != null)
                return source.LevelFilters.Select(Clone).ToList();

            var sources = new List<ScheduleItem> { source };
            var parameters = ScheduleBatchService.CollectParameters(doc, sources, new ScheduleParameterCatalog(doc), false);
            var result = new List<LevelScheduleFilter>();
            foreach (var common in ScheduleBatchService.ReadCommonFilters(doc, sources, parameters))
            {
                var legacy = source.Filters.Single(filter => filter.Index == common.SourceIndices[source.Schedule.UniqueId]);
                var rule = new LevelScheduleFilter
                {
                    IsExisting = true, Parameter = common.Parameter, Condition = common.Condition, Kind = common.Kind,
                    InitialParameterKey = common.InitialParameterKey, InitialKind = common.InitialKind,
                    Visibility = common.Visibility, InitialVisibility = common.InitialVisibility,
                    InitialSignature = common.InitialSignature, InitialFilterSignature = common.InitialFilterSignature,
                    OriginalIdentity = legacy.FieldIdentity, OriginalValue = common.Value, OriginalCondition = common.Condition,
                    IsLevelParameter = legacy.Original.IsElementIdValue
                        && doc.GetElement(legacy.Original.GetElementIdValue()) is Level,
                    Mode = legacy.Mode, Value = legacy.Mode == FilterModes.Keep ? common.Value : legacy.Pattern
                };
                rule.SourceIndices.Add(source.Schedule.UniqueId, legacy.Index);
                foreach (var pair in legacy.Overrides)
                    rule.Overrides.Add(pair.Key, pair.Value);
                result.Add(rule);
            }
            return result;
        }

        public static LevelScheduleFilter Clone(LevelScheduleFilter source)
        {
            var result = new LevelScheduleFilter
            {
                IsExisting = source.IsExisting, Parameter = source.Parameter, Condition = source.Condition,
                Kind = source.Kind, Value = source.Value, Enabled = source.Enabled, Deleted = source.Deleted,
                OriginalIdentity = source.OriginalIdentity, OriginalValue = source.OriginalValue,
                OriginalCondition = source.OriginalCondition, IsLevelParameter = source.IsLevelParameter,
                Mode = source.Mode, InitialSignature = source.InitialSignature,
                InitialParameterKey = source.InitialParameterKey, InitialKind = source.InitialKind,
                Visibility = source.Visibility, InitialVisibility = source.InitialVisibility,
                InitialFilterSignature = source.InitialFilterSignature
            };
            foreach (var pair in source.SourceIndices)
                result.SourceIndices.Add(pair.Key, pair.Value);
            foreach (var pair in source.Overrides)
                result.Overrides.Add(pair.Key, pair.Value);
            return result;
        }

        public static string ResolveText(LevelScheduleFilter rule, ScheduleLevelRow level, int index,
            bool includeManual = true, bool requireManual = false)
        {
            if (!rule.Enabled || (rule.IsExisting && !rule.HasChanges))
                return rule.IsExisting ? rule.OriginalValue : "";
            if (rule.Deleted || !rule.NeedsValue)
                return "";
            if (rule.Mode == FilterModes.Keep)
                return rule.OriginalValue;
            if (rule.Mode == FilterModes.Level)
            {
                if (!rule.IsLevelParameter || level.Level == null)
                    throw new InvalidOperationException("Для «Целевого уровня» нужен фильтр по уровню текущей модели.");
                return level.Level.Id.ToString();
            }

            string value;
            if (includeManual && rule.Overrides.TryGetValue(level.Key, out value))
                return value;
            if (rule.Mode == FilterModes.Manual)
            {
                if (requireManual)
                    throw new InvalidOperationException("Задайте значение вручную: " + level.Name + ", " + rule.Parameter?.Name);
                return rule.OriginalValue ?? "";
            }

            ScheduleMaskCatalog.CopyByLevels(allowOriginal: rule.IsExisting).Validate(rule.Value);
            return ScheduleNaming.Expand(rule.Value, level.Floor, level.Name, level.Elevation,
                level.Group, rule.OriginalValue, index);
        }

        internal static string ResolveValue(Document doc, LevelScheduleFilter rule, ScheduleLevelRow level, int index)
        {
            string text = ResolveText(rule, level, index, requireManual: true);
            long numeric;
            if (rule.NeedsValue && rule.Kind == "ID элемента" && !long.TryParse(text, out numeric))
            {
                // Старые настройки ссылочных фильтров могли хранить уникальное имя элемента.
                Element original = null;
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                int originalId;
                if (int.TryParse(rule.OriginalValue, out originalId))
#else
                long originalId;
                if (long.TryParse(rule.OriginalValue, out originalId))
#endif
                    original = doc.GetElement(new ElementId(originalId));
                var matches = new FilteredElementCollector(doc).WhereElementIsNotElementType()
                    .Concat(new FilteredElementCollector(doc).WhereElementIsElementType())
                    .Where(e => (original == null || e.GetType() == original.GetType()) && e.Name == text).Take(2).ToList();
                if (matches.Count != 1)
                    throw new InvalidOperationException("Укажите ID или уникальное имя элемента: " + rule.Parameter?.Name + " = " + text);
                return matches[0].Id.ToString();
            }
            return text;
        }

        public static bool HasChanges(ScheduleItem source)
        {
            return (source.LevelSorts?.Any(rule => rule.HasChanges) ?? false)
                || (source.LevelFilters == null ? source.Filters.Any(rule => rule.Mode != FilterModes.Keep)
                : source.LevelFilters.Any(rule => rule.HasChanges));
        }

        public static IEnumerable<string> Describe(ScheduleItem source, ScheduleLevelRow level, int index)
        {
            if (source.LevelFilters == null)
                return source.Filters.Where(rule => rule.Mode != FilterModes.Keep)
                    .Select(rule => source.Name + ": " + rule.FieldName + " → " + ScheduleService.ResolveText(rule, level, index));

            return source.LevelFilters.Where(rule => rule.HasChanges).Select(rule => source.Name + ": "
                + rule.Parameter?.Name + (rule.Deleted ? " → удалить" : " / " + rule.Condition + " → " + ResolveText(rule, level, index))
                + (rule.VisibilityChanged ? " / столбец: " + rule.Visibility : ""))
                .Concat((source.LevelSorts ?? new List<BatchScheduleSort>()).Where(rule => rule.HasChanges)
                    .Select(rule => source.Name + ": сортировка «" + rule.Parameter?.Name + "» / "
                        + (rule.Deleted ? "удалить" : rule.Position + " / " + rule.Direction + " / " + rule.Visibility)));
        }

        public static void CheckSource(Document doc, ScheduleItem source, IEnumerable<LevelScheduleFilter> rules)
        {
            foreach (var rule in rules.Where(rule => rule.IsExisting && rule.HasChanges))
            {
                int position;
                if (!rule.SourceIndices.TryGetValue(source.Schedule.UniqueId, out position)
                    || position < 0 || position >= source.Schedule.Definition.GetFilterCount())
                    throw new InvalidOperationException("Состав фильтров изменился. Откройте настройки заново: " + source.Name);

                var filter = source.Schedule.Definition.GetFilter(position);
                var field = source.Schedule.Definition.GetField(filter.FieldId);
                string identity = field.ParameterId + "/" + field.FieldId + "/" + filter.FilterType;
                string value = filter.IsElementIdValue ? filter.GetElementIdValue().ToString()
                    : ScheduleService.DisplayValue(doc, field, filter);
                if (identity != rule.OriginalIdentity || value != rule.OriginalValue)
                    throw new InvalidOperationException("Исходный фильтр изменился. Повторно выберите лист-эталон: " + source.Name);
            }
        }

        public static void Validate(Document doc, ScheduleItem source, IList<LevelScheduleFilter> rules,
            IEnumerable<ScheduleOperation> targets, IList<BatchScheduleSort> sorts = null)
        {
            CheckSource(doc, source, rules);
            sorts = sorts ?? source.LevelSorts;
            if (!rules.Any(rule => rule.HasChanges) && !(sorts?.Any(rule => rule.HasChanges) ?? false))
                return;

            // Пробуем тот же общий обработчик, который создаёт итоговую экспликацию.
            // Каждая копия и её шаблон откатываются, исходная спецификация не меняется.
            using (var transaction = new Transaction(doc, "KPLN: проверка фильтров по этажам"))
            {
                transaction.Start();
                try
                {
                    foreach (var target in targets)
                    {
                        using (var sub = new SubTransaction(doc))
                        {
                            sub.Start();
                            try
                            {
                                var copy = (ViewSchedule)doc.GetElement(source.Schedule.Duplicate(ViewDuplicateOption.Duplicate));
                                ScheduleBatchService.ApplySettings(doc, copy, rules, sorts, target.Level, target.Index,
                                    new List<string>(), source.Schedule.UniqueId);
                            }
                            catch (Exception ex)
                            {
                                throw new InvalidOperationException(target.Level.Name + ": " + ex.Message, ex);
                            }
                            finally { sub.RollBack(); }
                        }
                    }
                }
                finally { transaction.RollBack(); }
            }
        }

        public static XElement Save(ScheduleItem source)
        {
            if (source.LevelFilters == null)
                return null;

            return new XElement("LevelFilters", ScheduleSortingService.Save(source), source.LevelFilters.Select(rule => new XElement("Rule",
                new XAttribute("Existing", rule.IsExisting), new XAttribute("Enabled", rule.Enabled),
                new XAttribute("Deleted", rule.Deleted), new XAttribute("Mode", rule.Mode),
                new XAttribute("Condition", rule.Condition), new XAttribute("Kind", rule.Kind),
                new XAttribute("Visibility", rule.Visibility),
                new XAttribute("Pattern", rule.Value ?? ""), new XAttribute("Identity", rule.OriginalIdentity ?? ""),
                new XAttribute("Index", rule.SourceIndices.ContainsKey(source.Schedule.UniqueId) ? rule.SourceIndices[source.Schedule.UniqueId] : -1),
                new XAttribute("ParameterKey", rule.Parameter?.Key ?? ""), new XAttribute("ParameterName", rule.Parameter?.Name ?? ""),
                rule.Overrides.Select(pair => new XElement("Value", new XAttribute("Level", pair.Key), pair.Value ?? "")))));
        }

        public static void Load(Document doc, ScheduleItem source, XElement saved)
        {
            if (saved == null)
                return;

            // Совмещаем сохранённые настройки с актуальными исходными фильтрами по индексу и идентичности.
            var rules = ReadRules(doc, source);
            foreach (var item in saved.Elements("Rule"))
            {
                bool existing = (bool?)item.Attribute("Existing") ?? false;
                var rule = existing ? rules.FirstOrDefault(candidate =>
                    candidate.OriginalIdentity == (string)item.Attribute("Identity")
                    && candidate.SourceIndices[source.Schedule.UniqueId] == (int?)item.Attribute("Index")) : new LevelScheduleFilter();
                if (rule == null)
                    continue;

                if (!existing || ((string)item.Attribute("ParameterKey") != null && (string)item.Attribute("ParameterKey") != rule.Parameter?.Key))
                {
                    rule.Parameter = new ScheduleParameterChoice
                    {
                        Key = (string)item.Attribute("ParameterKey"), Name = (string)item.Attribute("ParameterName"),
                        ScheduleCount = 1, SuggestedKind = (string)item.Attribute("Kind")
                    };
                    if (!existing) rules.Add(rule);
                }
                rule.Visibility = (string)item.Attribute("Visibility") ?? rule.Visibility;
                rule.Enabled = (bool?)item.Attribute("Enabled") ?? true;
                rule.Deleted = (bool?)item.Attribute("Deleted") ?? false;
                rule.Condition = (string)item.Attribute("Condition") ?? rule.Condition;
                rule.Value = (string)item.Attribute("Pattern") ?? "";
                string mode = (string)item.Attribute("Mode");
                rule.Mode = rule.Modes.Contains(mode) ? mode : FilterModes.Pattern;
                rule.Overrides.Clear();
                foreach (var value in item.Elements("Value"))
                    rule.Overrides[(string)value.Attribute("Level")] = value.Value;
            }
            source.LevelFilters = rules;
            source.LevelSorts = ScheduleSortingService.Load(doc, source, saved.Element("Sorting"));
        }
    }
}
