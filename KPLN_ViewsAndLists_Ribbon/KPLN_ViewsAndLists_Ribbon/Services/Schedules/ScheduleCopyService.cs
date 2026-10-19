using Autodesk.Revit.DB;
using KPLN_ViewsAndLists_Ribbon.Common.Schedules;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace KPLN_ViewsAndLists_Ribbon.Services.Schedules
{
    internal static class ScheduleCopyService
    {
        public static bool Supports(ScheduleParameterChoice parameter)
        {
            string type = (parameter?.Key ?? "").Split('/').Last();
            return new[] { "Instance", "ElementType", "Room", "FromRoom", "ToRoom", "Space", "Material", "ProjectInfo" }.Contains(type);
        }

        public static List<ScheduleCopyValue> ReadValues(Document doc, ViewSchedule source, ScheduleParameterChoice parameter)
        {
            if (parameter == null) throw new InvalidOperationException("Выберите параметр копирования.");
            if (!Supports(parameter)) throw new InvalidOperationException("Для этого вида поля чтение значений не поддерживается. Выберите параметр элементов или типов.");
            if (source.Definition.IncludeLinkedFiles)
                throw new InvalidOperationException("Копирование по значениям пока доступно для спецификаций текущего проекта. В исходной спецификации включены элементы связей.");

            var values = new Dictionary<string, ScheduleCopyValue>(StringComparer.OrdinalIgnoreCase);
            using (var transaction = new Transaction(doc, "KPLN: чтение значений спецификации"))
            {
                transaction.Start();
                try
                {
                    // Только временная копия. Исходные фильтры и общий шаблон не меняются.
                    var probe = (ViewSchedule)doc.GetElement(source.Duplicate(ViewDuplicateOption.Duplicate));
                    probe.ViewTemplateId = ElementId.InvalidElementId;
                    var definition = probe.Definition;
                    ScheduleField field = FindOrAddField(doc, definition, parameter);
                    if (!definition.CanFilterByValue(field.FieldId))
                        throw new InvalidOperationException("Параметр «" + parameter.Name + "» не поддерживает фильтр «Равно».");

                    definition.SetFilters(definition.GetFilters().Where(filter =>
                        ScheduleBatchService.ParameterKey(doc, definition.GetField(filter.FieldId)) != parameter.Key).ToList());
                    doc.Regenerate();

                    Phase phase = doc.GetElement(probe.get_Parameter(BuiltInParameter.VIEW_PHASE)?.AsElementId() ?? ElementId.InvalidElementId) as Phase;
                    var types = new Dictionary<ElementId, Element>();
                    var parameterElement = doc.GetElement(field.ParameterId) as ParameterElement;
                    Definition parameterDefinition = parameterElement?.GetDefinition();
                    Func<Element, Parameter> getter = parameterDefinition != null
                        ? (Func<Element, Parameter>)(owner => owner.get_Parameter(parameterDefinition))
                        : owner => owner.get_Parameter((BuiltInParameter)int.Parse(field.ParameterId.ToString(), CultureInfo.InvariantCulture));
                    using (var collector = new FilteredElementCollector(doc, probe.Id).WhereElementIsNotElementType())
                    {
                        foreach (Element element in collector)
                        {
                            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                            foreach (Element owner in Owners(doc, element, field.FieldType, phase))
                            {
                                Parameter value = FindParameter(doc, owner, getter, field.FieldType == ScheduleFieldType.ElementType, types);
                                if (value == null || !value.HasValue) continue;
                                ScheduleCopyValue row = ReadValue(doc, value);
                                if (row == null || !seen.Add(row.Key)) continue;
                                ScheduleCopyValue existing;
                                if (values.TryGetValue(row.Key, out existing)) existing.ElementCount++;
                                else { row.ElementCount = 1; values.Add(row.Key, row); }
                            }
                        }
                    }
                }
                finally
                {
                    if (transaction.GetStatus() == TransactionStatus.Started) transaction.RollBack();
                }
            }
            return values.Values.OrderBy(value => value.Storage == StorageType.Double ? value.Number
                : value.Storage == StorageType.Integer ? value.Integer : 0)
                .ThenBy(value => value.Display, StringComparer.CurrentCultureIgnoreCase).ToList();
        }

        private static IEnumerable<Element> Owners(Document doc, Element element, ScheduleFieldType type, Phase phase)
        {
            var instance = element as FamilyInstance;
            switch (type)
            {
                case ScheduleFieldType.Room:
                    yield return instance == null ? null : phase == null ? instance.Room : instance.get_Room(phase); break;
                case ScheduleFieldType.FromRoom:
                    yield return instance == null ? null : phase == null ? instance.FromRoom : instance.get_FromRoom(phase); break;
                case ScheduleFieldType.ToRoom:
                    yield return instance == null ? null : phase == null ? instance.ToRoom : instance.get_ToRoom(phase); break;
                case ScheduleFieldType.Space:
                    yield return instance == null ? null : phase == null ? instance.Space : instance.get_Space(phase); break;
                case ScheduleFieldType.Material:
                    foreach (ElementId id in element.GetMaterialIds(false)) yield return doc.GetElement(id); break;
                case ScheduleFieldType.ProjectInfo:
                    yield return doc.ProjectInformation; break;
                default:
                    yield return element; break;
            }
        }

        private static Parameter FindParameter(Document doc, Element owner, Func<Element, Parameter> getter, bool typeOnly, Dictionary<ElementId, Element> types)
        {
            if (owner == null) return null;
            Parameter parameter = typeOnly ? null : getter(owner);
            if (parameter != null) return parameter;

            ElementId typeId = owner.GetTypeId();
            Element type;
            if (!types.TryGetValue(typeId, out type)) { type = doc.GetElement(typeId); types[typeId] = type; }
            return type == null ? null : getter(type);
        }

        private static ScheduleCopyValue ReadValue(Document doc, Parameter parameter)
        {
            var row = new ScheduleCopyValue { Storage = parameter.StorageType };
            switch (parameter.StorageType)
            {
                case StorageType.String:
                    row.Text = parameter.AsString();
                    if (string.IsNullOrWhiteSpace(row.Text)) return null;
                    row.Display = row.Text; break;
                case StorageType.Integer:
                    row.Integer = parameter.AsInteger();
                    row.Display = parameter.AsValueString() ?? row.Integer.ToString(CultureInfo.InvariantCulture); break;
                case StorageType.Double:
                    row.Number = parameter.AsDouble();
                    row.Display = parameter.AsValueString() ?? row.Number.ToString("R", CultureInfo.InvariantCulture); break;
                case StorageType.ElementId:
                    var id = parameter.AsElementId();
                    if (id == ElementId.InvalidElementId) return null;
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                    row.ElementIdValue = id.IntegerValue;
#else
                    row.ElementIdValue = id.Value;
#endif
                    row.Display = doc.GetElement(id)?.Name ?? parameter.AsValueString() ?? id.ToString(); break;
                default: return null;
            }
            return row;
        }

        private static ScheduleField FindOrAddField(Document doc, ScheduleDefinition definition, ScheduleParameterChoice parameter)
        {
            var field = definition.GetFieldOrder().Select(definition.GetField)
                .FirstOrDefault(candidate => ScheduleBatchService.ParameterKey(doc, candidate) == parameter.Key);
            if (field != null) return field;
            var available = definition.GetSchedulableFields().FirstOrDefault(candidate => ScheduleBatchService.ParameterKey(doc, candidate) == parameter.Key);
            if (available == null) throw new InvalidOperationException("Параметр больше не доступен в исходной спецификации: " + parameter.Name);
            field = definition.AddField(available);
            field.IsHidden = true;
            return field;
        }

        public static List<ScheduleFilter> WithCopyFilter(Document doc, ViewSchedule schedule, ScheduleOperation operation,
            IEnumerable<ScheduleFilter> currentFilters, List<string> messages)
        {
            ScheduleService.EnsureEditableFilters(doc, schedule, schedule.Name, messages);
            var definition = schedule.Definition;
            ScheduleField field = FindOrAddField(doc, definition, operation.CopyParameter);
            if (!definition.CanFilterByValue(field.FieldId))
                throw new InvalidOperationException("Нельзя применить «Равно» к параметру «" + operation.CopyParameter.Name + "».");
            var filters = currentFilters.Where(filter =>
                ScheduleBatchService.ParameterKey(doc, definition.GetField(filter.FieldId)) != operation.CopyParameter.Key).ToList();
            var value = operation.CopyValue;
            ScheduleFilter automatic;
            switch (value.Storage)
            {
                case StorageType.String: automatic = new ScheduleFilter(field.FieldId, ScheduleFilterType.Equal, value.Text); break;
                case StorageType.Integer: automatic = new ScheduleFilter(field.FieldId, ScheduleFilterType.Equal, value.Integer); break;
                case StorageType.Double: automatic = new ScheduleFilter(field.FieldId, ScheduleFilterType.Equal, value.Number); break;
                case StorageType.ElementId:
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                    automatic = new ScheduleFilter(field.FieldId, ScheduleFilterType.Equal, new ElementId(checked((int)value.ElementIdValue))); break;
#else
                    automatic = new ScheduleFilter(field.FieldId, ScheduleFilterType.Equal, new ElementId(value.ElementIdValue)); break;
#endif
                default: throw new InvalidOperationException("Неизвестный тип значения копирования.");
            }
            filters.Add(automatic);
            messages.Add(operation.CopyParameter.Name + " = «" + value.Display + "».");
            return filters;
        }
    }
}
