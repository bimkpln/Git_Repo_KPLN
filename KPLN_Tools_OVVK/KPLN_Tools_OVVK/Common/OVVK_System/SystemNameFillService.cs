using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools_OVVK.Common.OVVK_System
{
    /// <summary>
    /// Перенос ОВ/ВК-ветки systemParam: параметр назначения выбирается в менеджере.
    /// Все обращения к Revit выполняются из очереди команд.
    /// </summary>
    internal sealed class SystemNameFillService
    {
        private readonly Document _doc;
        private readonly string _parameterName;
        private readonly bool _useAbbreviation;
        private readonly Dictionary<ElementId, Element> _elements;
        private readonly Dictionary<ElementId, FillItem> _items = new Dictionary<ElementId, FillItem>();
        private readonly List<KeyValuePair<string, string>> _systemTypes;
        private readonly Dictionary<string, string> _resolvedNames = new Dictionary<string, string>(StringComparer.Ordinal);
        private readonly Dictionary<string, string> _resolutionErrors = new Dictionary<string, string>(StringComparer.Ordinal);
        private readonly HashSet<string> _blockedTypeParameters = new HashSet<string>(StringComparer.Ordinal);
        private readonly Dictionary<ElementId, string> _errors = new Dictionary<ElementId, string>();
        private readonly Action _checkCancellation;

        internal SystemNameFillService(Document doc, string parameterName, IEnumerable<Element> elements, Action checkCancellation = null)
        {
            _doc = doc;
            _parameterName = parameterName;
            _useAbbreviation = doc.Title.Contains("МТРС_") && doc.Title.Contains("_ВК_");
            _elements = elements.ToDictionary(element => element.Id);
            _checkCancellation = checkCancellation;
            HashSet<ElementId> usedTypeIds = GetUsedSystemTypeIds();
            _systemTypes = new List<KeyValuePair<string, string>>();
            foreach (BuiltInCategory category in new[] { BuiltInCategory.OST_PipingSystem, BuiltInCategory.OST_DuctSystem })
            {
                using (var types = new FilteredElementCollector(doc).OfCategory(category).WhereElementIsElementType())
                {
                    foreach (Element type in types)
                    {
                        _checkCancellation?.Invoke();
                        if (!usedTypeIds.Contains(type.Id))
                            continue;
                        _systemTypes.Add(new KeyValuePair<string, string>(
                            type.get_Parameter(BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM)?.AsString() ?? string.Empty,
                            type.Name));
                    }
                }
            }
        }


        internal IReadOnlyDictionary<ElementId, string> Errors => _errors;
        internal int ChangedCount { get; private set; }
        internal int UnchangedCount { get; private set; }
        internal int TypeParameterCount { get; private set; }
        internal string LogWarning { get; private set; }

        private HashSet<ElementId> GetUsedSystemTypeIds()
        {
            // Неиспользуемый дубликат сокращения не должен подмешивать имя постороннего типа.
            var usedIds = new HashSet<ElementId>();
            using (var systems = new FilteredElementCollector(_doc).OfClass(typeof(MechanicalSystem)))
            {
                foreach (MechanicalSystem system in systems)
                {
                    _checkCancellation?.Invoke();
                    ElementId typeId = system.GetTypeId();
                    if (!usedIds.Contains(typeId) && (system.Elements.Size > 0 || system.BaseEquipment != null
                        || system.DuctNetwork?.Size > 0))
                        usedIds.Add(typeId);
                }
            }
            using (var systems = new FilteredElementCollector(_doc).OfClass(typeof(PipingSystem)))
            {
                foreach (PipingSystem system in systems)
                {
                    _checkCancellation?.Invoke();
                    ElementId typeId = system.GetTypeId();
                    if (!usedIds.Contains(typeId) && (system.Elements.Size > 0 || system.BaseEquipment != null
                        || system.PipingNetwork?.Size > 0))
                        usedIds.Add(typeId);
                }
            }
            return usedIds;
        }

        internal void Prepare(Action<int, int, string> reportProgress = null)
        {
            int processed = 0;
            reportProgress?.Invoke(0, _elements.Count, "Определяю системы элементов");
            foreach (Element element in _elements.Values)
            {
                _checkCancellation?.Invoke();
                ResolveItem(element, new HashSet<ElementId>());
                reportProgress?.Invoke(++processed, _elements.Count, "Определяю системы элементов");
            }

            // Один параметр типа общий для всех экземпляров: неоднозначные записи пропускаем.
            reportProgress?.Invoke(0, 1, "Проверяю параметры типов и вложенные семейства");
            var parameterGroups = _items.Values.Where(item => item != null).ToLookup(item => item.Key);
            foreach (IGrouping<string, FillItem> group in parameterGroups)
            {
                _checkCancellation?.Invoke();
                if (_blockedTypeParameters.Contains(group.Key))
                {
                    foreach (FillItem item in group)
                        _errors[item.Element.Id] = "Не удалось определить систему другого экземпляра с тем же параметром типа. Используй параметр экземпляра или исправь замечания.";
                }
                else if (group.Select(item => item.Value).Distinct(StringComparer.Ordinal).Skip(1).Any())
                {
                    foreach (FillItem item in group)
                        _errors[item.Element.Id] = "Один параметр типа требует разных имён систем. Используй параметр экземпляра.";
                }
            }

            // Не переносим вычисленное значение родителя, если сам родитель не будет заполнен.
            var children = _items.Values.Where(item => item != null && item.ParentId != null)
                .ToLookup(item => item.ParentId);
            var failedParents = new Queue<ElementId>(_errors.Keys);
            var checkedTypeParameters = new HashSet<string>(StringComparer.Ordinal);
            while (failedParents.Count > 0)
            {
                _checkCancellation?.Invoke();
                ElementId parentId = failedParents.Dequeue();
                if (_items.TryGetValue(parentId, out FillItem failedItem) && failedItem?.Parameter.Element is ElementType
                    && checkedTypeParameters.Add(failedItem.Key))
                {
                    foreach (FillItem item in parameterGroups[failedItem.Key])
                    {
                        if (!_errors.ContainsKey(item.Element.Id))
                        {
                            _errors[item.Element.Id] = "Не удалось заполнить другой экземпляр с тем же параметром типа. Используй параметр экземпляра или исправь замечания.";
                            failedParents.Enqueue(item.Element.Id);
                        }
                    }
                }
                foreach (FillItem item in children[parentId])
                {
                    if (!_errors.ContainsKey(item.Element.Id))
                    {
                        _errors[item.Element.Id] = "Не удалось заполнить выбранный родительский элемент " + item.ParentId;
                        failedParents.Enqueue(item.Element.Id);
                    }
                }
            }

            reportProgress?.Invoke(1, 1, "Проверяю параметры типов и вложенные семейства");
        }

        internal void Write(Action<int, int, string> reportProgress = null)
        {
            // Вызывается внутри единой транзакции: неожиданный сбой записи отменяет всё заполнение.
            var groups = _items.Values.Where(item => item != null && !_errors.ContainsKey(item.Element.Id))
                .GroupBy(item => item.Key).ToArray();
            int processed = 0;
            reportProgress?.Invoke(0, groups.Length, "Заполняю параметр");
            foreach (IGrouping<string, FillItem> group in groups)
            {
                _checkCancellation?.Invoke();
                FillItem item = group.First();
                if (string.Equals(item.Parameter.AsString(), item.Value, StringComparison.Ordinal))
                {
                    UnchangedCount += group.Count();
                }
                else
                {
                    if (!item.Parameter.Set(item.Value))
                        throw new InvalidOperationException($"Не удалось записать «{_parameterName}» у элемента {item.Element.Id}.");

                    ChangedCount += group.Count();
                    if (item.Parameter.Element is ElementType)
                        TypeParameterCount++;
                }

                reportProgress?.Invoke(++processed, groups.Length, "Заполняю параметр");
            }

            _checkCancellation?.Invoke();
            if (ChangedCount > 0 || UnchangedCount > 0)
            {
                try
                {
                    WriteLog();
                }
                catch (Exception ex)
                {
                    LogWarning = "Не удалось записать журнал запуска: " + ex.Message;
                }
            }
        }


        private FillItem ResolveItem(Element element, HashSet<ElementId> resolving)
        {
            _checkCancellation?.Invoke();
            if (_items.TryGetValue(element.Id, out FillItem cached))
                return cached;

            Parameter target = null;
            try
            {
                if (!resolving.Add(element.Id))
                    throw new InvalidOperationException("Обнаружен цикл вложенности семейств.");

                target = GetParameter(element);
                if (target == null)
                    throw new InvalidOperationException($"Отсутствует параметр «{_parameterName}» у экземпляра и типа.");
                if (target.StorageType != StorageType.String)
                    throw new InvalidOperationException($"Параметр «{_parameterName}» должен быть текстовым.");

                string value;
                ElementId parentId = null;
                Element parent = (element as FamilyInstance)?.SuperComponent;
                if (parent != null)
                {
                    if (_elements.ContainsKey(parent.Id))
                    {
                        FillItem parentItem = ResolveItem(parent, resolving);
                        if (parentItem == null)
                            throw new InvalidOperationException("Не удалось определить систему выбранного родителя " + parent.Id);

                        value = parentItem.Value;
                        parentId = parent.Id;
                    }
                    else
                    {
                        Parameter parentParameter = GetParameter(parent);
                        if (parentParameter == null || parentParameter.StorageType != StorageType.String
                            || string.IsNullOrWhiteSpace(parentParameter.AsString()))
                            throw new InvalidOperationException($"У родителя {parent.Id} не заполнен «{_parameterName}». Заполни родительский элемент или запусти обработку всей модели.");

                        value = parentParameter.AsString();
                    }
                }
                else
                    value = ResolveSystemName(element);

                if (target.IsReadOnly && !string.Equals(target.AsString(), value, StringComparison.Ordinal))
                    throw new InvalidOperationException($"Параметр «{_parameterName}» доступен только для чтения.");

                var item = new FillItem(element, target, value, parentId);
                _items[element.Id] = item;
                return item;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                if (target?.Element is ElementType)
                    _blockedTypeParameters.Add(FillItem.GetKey(target));
                _errors[element.Id] = ex.Message;
                _items[element.Id] = null;
                return null;
            }
            finally
            {
                resolving.Remove(element.Id);
            }
        }

        private Parameter GetParameter(Element element)
        {
            IList<Parameter> parameters = element.GetParameters(_parameterName);
            if (parameters.Count == 0)
            {
                Element type = _doc.GetElement(element.GetTypeId());
                parameters = type?.GetParameters(_parameterName);
            }

            if (parameters == null || parameters.Count == 0)
                return null;
            if (parameters.Count != 1)
                throw new InvalidOperationException($"Найдено несколько параметров «{_parameterName}» у элемента {element.Id}.");

            return parameters[0];
        }

        private string ResolveSystemName(Element element)
        {
            Parameter source = element.get_Parameter(BuiltInParameter.RBS_DUCT_PIPE_SYSTEM_ABBREVIATION_PARAM);
            if (source != null && source.StorageType == StorageType.String)
            {
                string abbreviation = source.AsString();
                // Особое правило исходного скрипта МТРС ВК; параметр назначения остаётся выбранным в окне.
                return _useAbbreviation ? abbreviation ?? string.Empty : ResolveAbbreviation(abbreviation);
            }

            Parameter typeParameter = element.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM)
                ?? element.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);
            if (typeParameter != null && typeParameter.StorageType == StorageType.ElementId)
            {
                Element type = _doc.GetElement(typeParameter.AsElementId());
                if (type == null)
                    return _useAbbreviation ? string.Empty : "Нет системы";

                return _useAbbreviation
                    ? type.get_Parameter(BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM)?.AsString() ?? string.Empty
                    : type.Name;
            }

            throw new InvalidOperationException("Нет исходного параметра сокращения или типа системы.");
        }

        private string ResolveAbbreviation(string abbreviation)
        {
            // На больших моделях тысячи элементов имеют одинаковое сокращение.
            // Сопоставляем каждый набор сокращений с типами только один раз за запуск.
            string key = abbreviation ?? string.Empty;
            if (_resolvedNames.TryGetValue(key, out string value))
                return value;
            if (_resolutionErrors.TryGetValue(key, out string error))
                throw new InvalidOperationException(error);

            try
            {
                value = SystemNameResolver.Resolve(abbreviation, _systemTypes);
                _resolvedNames.Add(key, value);
                return value;
            }
            catch (InvalidOperationException ex)
            {
                _resolutionErrors.Add(key, ex.Message);
                throw;
            }
        }

        private void WriteLog()
        {
            // Сохраняем схему журнала Python, чтобы прежняя история запусков оставалась доступной.
            var schemaId = new Guid("be15305c-5249-4581-a4ca-01784efd8415");
            Schema schema = Schema.Lookup(schemaId);
            if (schema == null)
            {
                var builder = new SchemaBuilder(schemaId);
                builder.SetSchemaName("KPLN_SystemType");
                builder.SetReadAccessLevel(AccessLevel.Public);
                builder.SetWriteAccessLevel(AccessLevel.Public);
                builder.AddSimpleField("Last_Run", typeof(string));
                schema = builder.Finish();
            }

            Field field = schema.GetField("Last_Run");
            Entity previous = _doc.ProjectInformation.GetEntity(schema);
            string history = previous.IsValid() ? previous.Get<string>(field) : "---Начало логов---";
            string categories = string.Join("~", _elements.Values.Select(element => element.Category.Name).Distinct());
            string entry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} // {_doc.Application.Username} // "
                + $"Заполнение «{_parameterName}» на категории: {categories}; изменено: {ChangedCount}; замечаний: {_errors.Count}";

            var entity = new Entity(schema);
            entity.Set(field, history + Environment.NewLine + entry);
            _doc.ProjectInformation.SetEntity(entity);
        }


        private sealed class FillItem
        {
            internal FillItem(Element element, Parameter parameter, string value, ElementId parentId)
            {
                Element = element;
                Parameter = parameter;
                Value = value;
                ParentId = parentId;
                Key = GetKey(parameter);
            }

            internal static string GetKey(Parameter parameter) => parameter.Element.Id + ":" + parameter.Id;

            internal Element Element { get; }
            internal Parameter Parameter { get; }
            internal string Value { get; }
            internal ElementId ParentId { get; }
            internal string Key { get; }
        }
    }
}
