using Autodesk.Revit.DB;
using KPLN_Library_Schedules.Common.Schedules;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using FilterRule = KPLN_Library_Schedules.Common.Schedules.FilterRule;

namespace KPLN_Library_Schedules.Services.Schedules
{
    public static class ScheduleService
    {
        public static void ReadFilters(Document doc, ScheduleItem row)
        {
            ViewSchedule schedule = row.Schedule;
            row.Filters.Clear();
            row.Hint = schedule.ViewTemplateId == ElementId.InvalidElementId
                ? "Без шаблона" : "Шаблон: " + doc.GetElement(schedule.ViewTemplateId).Name;

            var definition = schedule.Definition;
            for (int index = 0; index < definition.GetFilterCount(); index++)
            {
                ScheduleFilter filter = definition.GetFilter(index);
                ScheduleField field = definition.GetField(filter.FieldId);
                row.Filters.Add(new FilterRule
                {
                    Index = index, Original = filter, FieldName = field.GetName(),
                    FieldIdentity = field.ParameterId + "/" + field.FieldId + "/" + filter.FilterType,
                    Value = ScheduleService.DisplayValue(doc, field, filter),
                    ValueType = filter.IsStringValue ? "Текст" : filter.IsIntegerValue ? "Целое число"
                        : filter.IsDoubleValue ? "Число в единицах проекта" : filter.IsElementIdValue ? "Элемент" : "Без значения"
                });
            }
        }

        public static string DisplayValue(Document doc, ScheduleField field, ScheduleFilter filter)
        {
            if (filter.IsStringValue) return filter.GetStringValue();
            if (filter.IsIntegerValue) return filter.GetIntegerValue().ToString(CultureInfo.InvariantCulture);
            if (filter.IsDoubleValue)
            {
#if Debug2020 || Revit2020
                return UnitFormatUtils.Format(doc.GetUnits(), field.UnitType, filter.GetDoubleValue(), false, false);
#else
                return UnitFormatUtils.Format(doc.GetUnits(), field.GetSpecTypeId(), filter.GetDoubleValue(), false);
#endif
            }
            if (filter.IsElementIdValue)
            {
                ElementId id = filter.GetElementIdValue();
                return (doc.GetElement(id)?.Name ?? "Элемент") + " [" + id + "]";
            }
            return "—";
        }

        public static string ResolveText(FilterRule rule, ScheduleLevelRow level, int index, bool includeManual = true)
        {
            string value;

            if (rule.Mode == FilterModes.Keep) return rule.Value;
            if (rule.Mode == FilterModes.Level)
            {
                if (level.Level == null) throw new InvalidOperationException("Для режима «Целевой уровень» выберите уровень.");
                return level.Name;
            }
            if (includeManual && rule.Overrides.TryGetValue(level.Key, out value)) return value;
            if (rule.Mode == FilterModes.Manual) return rule.Value;
            return ScheduleNaming.Expand(rule.Pattern, level.Floor, level.Name, level.Elevation, level.Group, rule.Value, index);
        }

        public static ScheduleFilter ResolveFilter(Document doc, ViewSchedule source, FilterRule rule, ScheduleLevelRow level, int index)
        {
            ScheduleFilter original = source.Definition.GetFilter(rule.Index);
            if (rule.Mode == FilterModes.Keep) return original;
            ScheduleField field = source.Definition.GetField(original.FieldId);

            if (original.IsNullValue)
                throw new InvalidOperationException("Условие «" + rule.FieldName + "» не имеет заменяемого значения.");

            if (rule.Mode == FilterModes.Level)
            {
                if (level.Level == null || !original.IsElementIdValue || !(doc.GetElement(original.GetElementIdValue()) is Level))
                    throw new InvalidOperationException("Режим «Целевой уровень» доступен для фильтра по уровню текущей модели: " + rule.FieldName);
                return new ScheduleFilter(original.FieldId, original.FilterType, level.Level.Id);
            }

            if (rule.Mode == FilterModes.Manual && !rule.Overrides.ContainsKey(level.Key))
                throw new InvalidOperationException("Задайте значение вручную: " + level.Name + ", " + rule.FieldName);

            string text = ResolveText(rule, level, index);
            if (original.IsStringValue) return new ScheduleFilter(original.FieldId, original.FilterType, text);
            int integer;
            if (original.IsIntegerValue && int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out integer))
                return new ScheduleFilter(original.FieldId, original.FilterType, integer);
            if (original.IsDoubleValue)
            {
                double number;
#if Debug2020 || Revit2020
                bool parsed = UnitFormatUtils.TryParse(doc.GetUnits(), field.UnitType, text, out number);
#else
                bool parsed = UnitFormatUtils.TryParse(doc.GetUnits(), field.GetSpecTypeId(), text, out number);
#endif
                if (parsed) return new ScheduleFilter(original.FieldId, original.FilterType, number);
            }
            if (original.IsElementIdValue)
            {
                // Для ссылочных полей выбирается реальный элемент того же класса; неоднозначные имена запрещены.
                Element oldElement = doc.GetElement(original.GetElementIdValue());
                if (oldElement == null) throw new InvalidOperationException("Не удалось определить тип ссылочного фильтра: " + rule.FieldName);
                var candidates = new FilteredElementCollector(doc);
                if (oldElement is ElementType) candidates.WhereElementIsElementType();
                else candidates.WhereElementIsNotElementType();
                var matches = candidates.Where(e => e.GetType() == oldElement.GetType() && e.Name == text).ToList();
                if (matches.Count == 1) return new ScheduleFilter(original.FieldId, original.FilterType, matches[0].Id);
                throw new InvalidOperationException("Для «" + rule.FieldName + "» нужно уникальное имя элемента: " + text);
            }
            throw new InvalidOperationException("Значение «" + text + "» не соответствует типу «" + rule.ValueType + "» поля «" + rule.FieldName + "».");
        }

        public static ViewSchedule Create(Document doc, ScheduleItem row, ScheduleLevelRow level, int index, List<string> warnings, string name = null)
        {
            var source = row.Schedule;
            if (!row.Copy) return source;
            if (!source.CanViewBeDuplicated(ViewDuplicateOption.Duplicate))
                throw new InvalidOperationException("Спецификация не поддерживает копирование: " + source.Name);

            var copy = (ViewSchedule)doc.GetElement(source.Duplicate(ViewDuplicateOption.Duplicate));
            copy.Name = name ?? ScheduleSupport.UniqueViewName(doc, source.Name + " — " + level.Name);
            if (row.LevelFilters == null)
                ApplyFilters(doc, source, copy, row.Filters, level, index, warnings);
            else
            {
                ScheduleLevelFilterService.CheckSource(doc, row, row.LevelFilters);
                ScheduleBatchService.ApplySettings(doc, copy, row.LevelFilters, row.LevelSorts, level, index, warnings, source.UniqueId);
            }
            return copy;
        }

        public static void ApplyFilters(Document doc, ViewSchedule source, ViewSchedule copy, IEnumerable<FilterRule> rules, ScheduleLevelRow level, int index, List<string> warnings)
        {
            var ruleList = rules.ToList();
            var changing = ruleList.Where(r => r.Mode != FilterModes.Keep).ToList();
            // Дубликат уже содержит неизменённые фильтры. Не перестраиваем таблицу
            // до применения пакетных условий, если индивидуальных правок нет.
            if (changing.Count == 0) return;
            var expectedFilters = ruleList.ToDictionary(r => r.Index, r => ResolveFilter(doc, source, r, level, index));
            EnsureEditableFilters(doc, copy, level.Name, warnings);

            // Индексы и поля берём из копии; остальные строки фильтров не трогаем.
            foreach (FilterRule rule in changing)
            {
                ScheduleFilter resolved = expectedFilters[rule.Index];
                ScheduleFilter existing = copy.Definition.GetFilter(rule.Index);
                resolved.FieldId = existing.FieldId;
                copy.Definition.SetFilter(rule.Index, resolved);
                warnings.Add(rule.FieldName + ": " + DisplayValue(doc, copy.Definition.GetField(existing.FieldId), resolved) + ".");
            }
            doc.Regenerate();

            if (copy.Definition.GetFilterCount() != source.Definition.GetFilterCount())
                throw new InvalidOperationException("Изменилось число фильтров спецификации: " + source.Name);
            foreach (FilterRule rule in ruleList)
            {
                ScheduleFilter expected = expectedFilters[rule.Index];
                if (!SameValue(expected, copy.Definition.GetFilter(rule.Index)))
                    throw new InvalidOperationException("Revit не сохранил фильтр: " + rule.FieldName);
            }

            var body = copy.GetTableData().GetSectionData(SectionType.Body);
            if (body.NumberOfRows <= 1) warnings.Add("Проверьте пустую экспликацию: " + copy.Name);
        }

        public static void EnsureEditableFilters(Document doc, ViewSchedule schedule, string suffix, List<string> warnings)
        {
            if (ScheduleSupport.TemplateControls(doc, schedule, BuiltInParameter.SCHEDULE_FILTER_PARAM)
                || ScheduleSupport.TemplateControls(doc, schedule, BuiltInParameter.SCHEDULE_FIELDS_PARAM))
            {
                // Разрешаем изменение фильтров только в отдельном шаблоне копии.
                var template = (View)doc.GetElement(schedule.ViewTemplateId);
                var ownTemplate = (View)doc.GetElement(template.Duplicate(ViewDuplicateOption.Duplicate));
                ownTemplate.Name = ScheduleSupport.UniqueViewName(doc, template.Name + " — " + suffix);
                var free = ownTemplate.GetNonControlledTemplateParameterIds().ToList();
                var parameter = new ElementId(BuiltInParameter.SCHEDULE_FILTER_PARAM);
                if (!free.Contains(parameter)) free.Add(parameter);
                var fields = new ElementId(BuiltInParameter.SCHEDULE_FIELDS_PARAM);
                if (!free.Contains(fields)) free.Add(fields);
                ownTemplate.SetNonControlledTemplateParameterIds(free);
                schedule.ViewTemplateId = ownTemplate.Id;
                warnings.Add("Создан индивидуальный шаблон спецификации: " + ownTemplate.Name);
            }

        }

        public static bool SameValue(ScheduleFilter a, ScheduleFilter b)
        {
            if (a.FilterType != b.FilterType) return false;
            if (a.IsStringValue) return b.IsStringValue && a.GetStringValue() == b.GetStringValue();
            if (a.IsIntegerValue) return b.IsIntegerValue && a.GetIntegerValue() == b.GetIntegerValue();
            if (a.IsDoubleValue) return b.IsDoubleValue && Math.Abs(a.GetDoubleValue() - b.GetDoubleValue()) < 1e-9;
            if (a.IsElementIdValue) return b.IsElementIdValue && a.GetElementIdValue() == b.GetElementIdValue();
            return b.IsNullValue;
        }
    }
}
