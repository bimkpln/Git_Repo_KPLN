using Autodesk.Revit.DB;
using KPLN_RevitMcpBridge.Server;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace KPLN_RevitMcpBridge.Services
{
    internal static class FamilyTypeService
    {
        private sealed class ValueUpdate
        {
            public FamilyParameter Parameter;
            public object ExpectedValue;
            public object Value;
            public object Describe() => new
            {
                parameter_id = RevitService.IdValue(Parameter.Id).ToString(CultureInfo.InvariantCulture),
                parameter_name = Parameter.Definition.Name,
                value = Value,
                outcome = "planned"
            };
            public object DescribeChange(string outcome) => new
            {
                parameter_id = RevitService.IdValue(Parameter.Id).ToString(CultureInfo.InvariantCulture),
                parameter_name = Parameter.Definition.Name,
                expected_value = ExpectedValue,
                value = Value,
                outcome
            };
        }

        private sealed class TypeCreate
        {
            public string Name;
            public List<ValueUpdate> Values = new List<ValueUpdate>();
            public FamilyType Created;
            public object Describe(string outcome) => new
            {
                name = Name,
                outcome,
                values = Values.Select(v => v.Describe()).ToArray()
            };
        }

        private sealed class TypeUpdate
        {
            public string Name;
            public FamilyType Type;
            public List<ValueUpdate> Values = new List<ValueUpdate>();
            public object Describe(string outcome) => new
            {
                name = Name,
                outcome,
                values = Values.Select(v => v.DescribeChange(outcome)).ToArray()
            };
        }

        private static object Raw(FamilyType type, FamilyParameter parameter)
        {
            if (!type.HasValue(parameter)) return null;
            switch (parameter.StorageType)
            {
                case StorageType.String: return type.AsString(parameter);
                case StorageType.Integer: return type.AsInteger(parameter);
                case StorageType.Double: return type.AsDouble(parameter);
                case StorageType.ElementId: return RevitService.IdValue(type.AsElementId(parameter)).ToString(CultureInfo.InvariantCulture);
                default: return null;
            }
        }

        private static object Typed(FamilyParameter parameter, object value, string units)
        {
            if (value == null) throw new BridgeException("invalid_value", "Значение параметра типоразмера не может быть null.");
            switch (parameter.StorageType)
            {
                case StorageType.String:
                    if (!(value is string)) throw new BridgeException("invalid_value", "Для " + parameter.Definition.Name + " требуется строка.");
                    return value;
                case StorageType.Integer:
                    long integer = Json.Integer(value, "value");
                    if (integer < int.MinValue || integer > int.MaxValue) throw new BridgeException("invalid_value", "Целое число вне диапазона Int32.");
                    return (int)integer;
                case StorageType.Double:
                    double number = Json.Double(value);
                    if (units != "revit_internal") throw new BridgeException("invalid_units", "Для Double явно задайте value_units=revit_internal.");
                    return number;
                case StorageType.ElementId:
                    return Json.Integer(value, "value").ToString(CultureInfo.InvariantCulture);
                default:
                    throw new BridgeException("unsupported_parameter", "Тип параметра не поддерживается: " + parameter.Definition.Name);
            }
        }

        private static void Set(FamilyManager manager, ValueUpdate update)
        {
            switch (update.Parameter.StorageType)
            {
                case StorageType.String: manager.Set(update.Parameter, (string)update.Value); break;
                case StorageType.Integer: manager.Set(update.Parameter, (int)update.Value); break;
                case StorageType.Double: manager.Set(update.Parameter, (double)update.Value); break;
                case StorageType.ElementId: manager.Set(update.Parameter, RevitService.Id(Json.Integer(update.Value, "value"))); break;
                default: throw new BridgeException("unsupported_parameter", "Тип параметра не поддерживается.");
            }
        }

        internal static object Create(Document doc, Dictionary<string, object> input)
        {
            if (!doc.IsFamilyDocument) throw new BridgeException("not_family_document", "Команда требует открытого редактора семейства.");
            if (doc.IsReadOnly || doc.IsLinked || doc.IsModifiable)
                throw new BridgeException("document_not_writable", "Документ недоступен для отдельной транзакции.", 409);
            bool dryRun = input.Flag("dry_run", true);
            string baseTypeName = input.Text("base_type_name", true);
            var manager = doc.FamilyManager;
            var existing = manager.Types.Cast<FamilyType>().ToDictionary(t => t.Name, StringComparer.Ordinal);
            FamilyType baseType;
            if (!existing.TryGetValue(baseTypeName, out baseType))
                throw new BridgeException("not_found", "Базовый типоразмер не найден: " + baseTypeName, 404);
            var parameters = manager.Parameters.Cast<FamilyParameter>().ToDictionary(p => RevitService.IdValue(p.Id));
            var plannedNames = new HashSet<string>(StringComparer.Ordinal);
            var creates = new List<TypeCreate>();
            foreach (var item in input.List("types", 20))
            {
                var row = item as Dictionary<string, object>;
                if (row == null) throw new BridgeException("invalid_input", "Каждый типоразмер должен быть объектом.");
                string name = row.Text("name", true);
                if (!plannedNames.Add(name)) throw new BridgeException("duplicate_type", "Типоразмер указан дважды: " + name);
                if (existing.ContainsKey(name)) throw new BridgeException("type_exists", "Типоразмер уже существует: " + name, 409);
                var create = new TypeCreate { Name = name };
                var seenParameters = new HashSet<long>();
                foreach (var valueItem in row.List("values", 200))
                {
                    var valueRow = valueItem as Dictionary<string, object>;
                    if (valueRow == null) throw new BridgeException("invalid_input", "Значение типоразмера должно быть объектом.");
                    long parameterId = Json.Integer(valueRow.Get("parameter_id"), "parameter_id");
                    if (!seenParameters.Add(parameterId)) throw new BridgeException("duplicate_update", "Один параметр типоразмера указан дважды.");
                    FamilyParameter parameter;
                    if (!parameters.TryGetValue(parameterId, out parameter)) throw new BridgeException("not_found", "Параметр FamilyManager не найден: " + parameterId, 404);
                    if (parameter.IsReadOnly || !string.IsNullOrEmpty(parameter.Formula))
                        throw new BridgeException("read_only_parameter", "Параметр нельзя изменять: " + parameter.Definition.Name);
                    string units = valueRow.Text("value_units");
                    object typed = Typed(parameter, valueRow.Get("value"), units);
                    if (parameter.StorageType == StorageType.ElementId)
                    {
                        var target = RevitService.Id(Json.Integer(typed, "value"));
                        if (RevitService.IdValue(target) >= 0 && doc.GetElement(target) == null)
                            throw new BridgeException("not_found", "Ссылка на ElementId не существует: " + typed, 404);
                    }
                    create.Values.Add(new ValueUpdate { Parameter = parameter, Value = typed });
                }
                creates.Add(create);
            }
            if (dryRun) return new
            {
                dry_run = true,
                validation = "family document, revision, absent type names, writable parameters and values; Revit constraints checked on commit",
                base_type_name = baseTypeName,
                types = creates.Select(c => c.Describe("planned")).ToArray()
            };

            var failures = new FailureCollector();
            using (var group = new TransactionGroup(doc, "KPLN: типоразмеры из Codex"))
            using (var transaction = new Transaction(doc, "KPLN: типоразмеры из Codex"))
            {
                if (group.Start() != TransactionStatus.Started) throw new BridgeException("transaction_failed", "Не удалось начать группу транзакций.");
                if (transaction.Start() != TransactionStatus.Started) throw new BridgeException("transaction_failed", "Не удалось начать транзакцию.");
                transaction.SetFailureHandlingOptions(transaction.GetFailureHandlingOptions().SetFailuresPreprocessor(failures).SetClearAfterRollback(true).SetForcedModalHandling(true));
                try
                {
                    foreach (var create in creates)
                    {
                        manager.CurrentType = baseType;
                        create.Created = manager.NewType(create.Name);
                        foreach (var update in create.Values) Set(manager, update);
                        doc.Regenerate();
                        foreach (var update in create.Values)
                            if (!object.Equals(Raw(create.Created, update.Parameter), update.Value))
                                throw new BridgeException("verification_failed", "Revit изменил параметр " + update.Parameter.Definition.Name + " типоразмера " + create.Name);
                    }
                    if (transaction.Commit() != TransactionStatus.Committed)
                        throw new BridgeException("transaction_rolled_back", "Revit откатил типоразмеры: " + string.Join("; ", failures.Messages), 409);
                    foreach (var create in creates)
                        foreach (var update in create.Values)
                            if (!object.Equals(Raw(create.Created, update.Parameter), update.Value))
                                throw new BridgeException("verification_failed", "Значение изменилось при Commit; пакет отменён.");
                    if (group.Assimilate() != TransactionStatus.Committed)
                        throw new BridgeException("transaction_failed", "Не удалось завершить группу транзакций.");
                }
                catch
                {
                    if (transaction.GetStatus() == TransactionStatus.Started) transaction.RollBack();
                    if (group.GetStatus() == TransactionStatus.Started) group.RollBack();
                    throw;
                }
            }
            return new
            {
                dry_run = false,
                base_type_name = baseTypeName,
                types = creates.Select(c => c.Describe("created")).ToArray(),
                warnings = failures.Messages.ToArray(),
                saved_to_disk = false
            };
        }

        internal static object Update(Document doc, Dictionary<string, object> input)
        {
            if (!doc.IsFamilyDocument) throw new BridgeException("not_family_document", "Команда требует открытого редактора семейства.");
            if (doc.IsReadOnly || doc.IsLinked || doc.IsModifiable)
                throw new BridgeException("document_not_writable", "Документ недоступен для отдельной транзакции.", 409);
            bool dryRun = input.Flag("dry_run", true);
            var manager = doc.FamilyManager;
            var existing = manager.Types.Cast<FamilyType>().ToDictionary(t => t.Name, StringComparer.Ordinal);
            var parameters = manager.Parameters.Cast<FamilyParameter>().ToDictionary(p => RevitService.IdValue(p.Id));
            var plannedNames = new HashSet<string>(StringComparer.Ordinal);
            var updates = new List<TypeUpdate>();

            foreach (var item in input.List("types", 20))
            {
                var row = item as Dictionary<string, object>;
                if (row == null) throw new BridgeException("invalid_input", "Каждый типоразмер должен быть объектом.");
                string name = row.Text("name", true);
                if (!plannedNames.Add(name)) throw new BridgeException("duplicate_type", "Типоразмер указан дважды: " + name);
                FamilyType type;
                if (!existing.TryGetValue(name, out type))
                    throw new BridgeException("not_found", "Типоразмер не найден: " + name, 404);

                var update = new TypeUpdate { Name = name, Type = type };
                var seenParameters = new HashSet<long>();
                foreach (var valueItem in row.List("values", 200))
                {
                    var valueRow = valueItem as Dictionary<string, object>;
                    if (valueRow == null) throw new BridgeException("invalid_input", "Значение типоразмера должно быть объектом.");
                    long parameterId = Json.Integer(valueRow.Get("parameter_id"), "parameter_id");
                    if (!seenParameters.Add(parameterId)) throw new BridgeException("duplicate_update", "Один параметр типоразмера указан дважды.");
                    FamilyParameter parameter;
                    if (!parameters.TryGetValue(parameterId, out parameter))
                        throw new BridgeException("not_found", "Параметр FamilyManager не найден: " + parameterId, 404);
                    if (parameter.IsReadOnly || !string.IsNullOrEmpty(parameter.Formula))
                        throw new BridgeException("read_only_parameter", "Параметр нельзя изменять: " + parameter.Definition.Name);
                    if (!valueRow.ContainsKey("expected_value"))
                        throw new BridgeException("invalid_input", "Не задано expected_value для параметра " + parameter.Definition.Name);

                    string units = valueRow.Text("value_units");
                    object expected = valueRow.Get("expected_value") == null ? null : Typed(parameter, valueRow.Get("expected_value"), units);
                    object typed = Typed(parameter, valueRow.Get("value"), units);
                    object actual = Raw(type, parameter);
                    if (!object.Equals(actual, expected))
                        throw new BridgeException("stale_value", "Значение параметра " + parameter.Definition.Name + " типоразмера " + name + " изменилось; перечитайте FamilyManager.", 409);
                    if (parameter.StorageType == StorageType.ElementId)
                    {
                        var target = RevitService.Id(Json.Integer(typed, "value"));
                        if (RevitService.IdValue(target) >= 0 && doc.GetElement(target) == null)
                            throw new BridgeException("not_found", "Ссылка на ElementId не существует: " + typed, 404);
                    }
                    update.Values.Add(new ValueUpdate { Parameter = parameter, ExpectedValue = expected, Value = typed });
                }
                updates.Add(update);
            }

            if (dryRun) return new
            {
                dry_run = true,
                validation = "family document, revision, existing type names, expected values, writable parameters and values; Revit constraints checked on commit",
                types = updates.Select(u => u.Describe("planned")).ToArray()
            };

            var originalType = manager.CurrentType;
            var failures = new FailureCollector();
            using (var group = new TransactionGroup(doc, "KPLN: параметры типоразмеров из Codex"))
            using (var transaction = new Transaction(doc, "KPLN: параметры типоразмеров из Codex"))
            {
                if (group.Start() != TransactionStatus.Started) throw new BridgeException("transaction_failed", "Не удалось начать группу транзакций.");
                if (transaction.Start() != TransactionStatus.Started) throw new BridgeException("transaction_failed", "Не удалось начать транзакцию.");
                transaction.SetFailureHandlingOptions(transaction.GetFailureHandlingOptions().SetFailuresPreprocessor(failures).SetClearAfterRollback(true).SetForcedModalHandling(true));
                try
                {
                    foreach (var update in updates)
                    {
                        manager.CurrentType = update.Type;
                        foreach (var value in update.Values) Set(manager, value);
                        doc.Regenerate();
                        foreach (var value in update.Values)
                            if (!object.Equals(Raw(update.Type, value.Parameter), value.Value))
                                throw new BridgeException("verification_failed", "Revit изменил параметр " + value.Parameter.Definition.Name + " типоразмера " + update.Name);
                    }
                    if (originalType != null) manager.CurrentType = originalType;
                    if (transaction.Commit() != TransactionStatus.Committed)
                        throw new BridgeException("transaction_rolled_back", "Revit откатил изменения типоразмеров: " + string.Join("; ", failures.Messages), 409);
                    foreach (var update in updates)
                        foreach (var value in update.Values)
                            if (!object.Equals(Raw(update.Type, value.Parameter), value.Value))
                                throw new BridgeException("verification_failed", "Значение изменилось при Commit; пакет отменён.");
                    if (group.Assimilate() != TransactionStatus.Committed)
                        throw new BridgeException("transaction_failed", "Не удалось завершить группу транзакций.");
                }
                catch
                {
                    if (transaction.GetStatus() == TransactionStatus.Started) transaction.RollBack();
                    if (group.GetStatus() == TransactionStatus.Started) group.RollBack();
                    throw;
                }
            }
            return new
            {
                dry_run = false,
                types = updates.Select(u => u.Describe("updated")).ToArray(),
                warnings = failures.Messages.ToArray(),
                saved_to_disk = false
            };
        }

        private sealed class FailureCollector : IFailuresPreprocessor
        {
            public readonly List<string> Messages = new List<string>();
            public FailureProcessingResult PreprocessFailures(FailuresAccessor accessor)
            {
                var messages = accessor.GetFailureMessages();
                foreach (var message in messages) Messages.Add(message.GetDescriptionText());
                return messages.Count > 0 ? FailureProcessingResult.ProceedWithRollBack : FailureProcessingResult.Continue;
            }
        }
    }
}
