using Autodesk.Revit.DB;
using KPLN_RevitMcpBridge.Server;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace KPLN_RevitMcpBridge.Services
{
    internal static class ParameterService
    {
        internal static object Raw(Parameter p)
        {
            if (!p.HasValue) return null;
            switch (p.StorageType)
            {
                case StorageType.String: return p.AsString();
                case StorageType.Integer: return p.AsInteger();
                case StorageType.Double: return p.AsDouble();
                case StorageType.ElementId: return RevitService.IdValue(p.AsElementId()).ToString(CultureInfo.InvariantCulture);
                default: return null;
            }
        }
        private static string DataType(Parameter p)
        {
#if Debug2020 || Revit2020
            return p.Definition.ParameterType.ToString();
#else
            return p.Definition.GetDataType().TypeId;
#endif
        }
        internal static object[] Read(Element e) => e.Parameters.Cast<Parameter>().OrderBy(p => RevitService.IdValue(p.Id)).Select(p => (object)new
        {
            id = RevitService.IdValue(p.Id).ToString(CultureInfo.InvariantCulture), name = p.Definition.Name,
            storage_type = p.StorageType.ToString(), data_type = DataType(p), value = Raw(p), display_value = p.AsValueString(),
            has_value = p.HasValue, is_read_only = p.IsReadOnly, shared_guid = p.IsShared ? p.GUID.ToString() : null,
            value_units = p.StorageType == StorageType.Double ? "revit_internal" : null
        }).ToArray();

        private sealed class Update
        {
            public Element Element;
            public Parameter Parameter;
            public object Before;
            public object After;
            public object Describe(string outcome) => new { element_unique_id = Element.UniqueId, element_id = RevitService.IdValue(Element.Id).ToString(CultureInfo.InvariantCulture),
                parameter_id = RevitService.IdValue(Parameter.Id).ToString(CultureInfo.InvariantCulture), parameter_name = Parameter.Definition.Name,
                before = Before, after = After, outcome };
        }
        private static object Typed(Parameter p, object value, bool allowNull)
        {
            if (value == null)
            {
                if (allowNull) return null;
                throw new BridgeException("invalid_value", "Новое значение не может быть null; для пустой строки используйте \"\".");
            }
            switch (p.StorageType)
            {
                case StorageType.String:
                    if (!(value is string)) throw new BridgeException("invalid_value", "Требуется строка.");
                    return value;
                case StorageType.Integer:
                    long integer = Json.Integer(value, "value");
                    if (integer < int.MinValue || integer > int.MaxValue) throw new BridgeException("invalid_value", "Целое число вне диапазона Int32.");
                    return (int)integer;
                case StorageType.Double:
                    return Json.Double(value);
                case StorageType.ElementId:
                    return Json.Integer(value, "value").ToString(CultureInfo.InvariantCulture);
                default: throw new BridgeException("unsupported_parameter", "Тип параметра не поддерживается.");
            }
        }
        private static bool Equal(object a, object b) => object.Equals(a, b);

        internal static object Set(Document doc, Dictionary<string, object> input)
        {
            if (doc.IsReadOnly || doc.IsLinked || doc.IsModifiable) throw new BridgeException("document_not_writable", "Документ недоступен для отдельной транзакции.", 409);
            if (doc.IsFamilyDocument) throw new BridgeException("family_write_not_supported", "В редакторе семейства пока доступно только чтение параметров FamilyManager.");
            bool dryRun = input.Flag("dry_run", true);
            var updates = new List<Update>(); var seen = new HashSet<string>();
            foreach (var item in input.List("updates"))
            {
                var row = item as Dictionary<string, object>;
                if (row == null) throw new BridgeException("invalid_input", "Каждое изменение должно быть объектом.");
                var uniqueId = row.Text("element_unique_id", true);
                long parameterId = Json.Integer(row.Get("parameter_id"), "parameter_id");
                if (!seen.Add(uniqueId + "/" + parameterId)) throw new BridgeException("duplicate_update", "Один параметр указан дважды.");
                var element = doc.GetElement(uniqueId);
                if (element == null) throw new BridgeException("not_found", "Элемент не найден: " + uniqueId, 404);
                var parameter = element.Parameters.Cast<Parameter>().SingleOrDefault(p => RevitService.IdValue(p.Id) == parameterId);
                if (parameter == null) throw new BridgeException("not_found", "Параметр не найден по Id.", 404);
                if (parameter.IsReadOnly) throw new BridgeException("read_only_parameter", "Параметр доступен только для чтения: " + parameter.Definition.Name);
                if (!row.ContainsKey("expected_value")) throw new BridgeException("invalid_input", "Требуется expected_value из get_elements.");
                object before = Raw(parameter), expected = Typed(parameter, row.Get("expected_value"), true), after = Typed(parameter, row.Get("value"), false);
                if (!Equal(before, expected)) throw new BridgeException("stale_parameter", "Значение параметра изменилось: " + parameter.Definition.Name, 409);
                if (parameter.StorageType == StorageType.Double && row.Text("value_units", true) != "revit_internal")
                    throw new BridgeException("invalid_units", "Для Double явно задайте value_units=revit_internal.");
                if (parameter.StorageType == StorageType.ElementId)
                {
                    var target = RevitService.Id(Json.Integer(after, "value"));
                    if (RevitService.IdValue(target) >= 0 && doc.GetElement(target) == null) throw new BridgeException("not_found", "Ссылка на ElementId не существует.");
                }
                updates.Add(new Update { Element = element, Parameter = parameter, Before = before, After = after });
            }
            if (dryRun) return new { dry_run = true, validation = "inputs_and_current_values; Revit constraints checked on commit", updates = updates.Select(u => u.Describe(Equal(u.Before, u.After) ? "unchanged" : "planned")).ToArray() };
            var failures = new FailureCollector();
            var changed = updates.Where(u => !Equal(u.Before, u.After)).ToArray();
            if (changed.Length > 0)
            {
                using (var group = new TransactionGroup(doc, "KPLN: параметры из Codex"))
                using (var transaction = new Transaction(doc, "KPLN: параметры из Codex"))
                {
                    if (group.Start() != TransactionStatus.Started) throw new BridgeException("transaction_failed", "Не удалось начать группу транзакций.");
                    if (transaction.Start() != TransactionStatus.Started) throw new BridgeException("transaction_failed", "Не удалось начать транзакцию.");
                    transaction.SetFailureHandlingOptions(transaction.GetFailureHandlingOptions().SetFailuresPreprocessor(failures).SetClearAfterRollback(true).SetForcedModalHandling(true));
                    try
                    {
                        foreach (var update in changed)
                        {
                            var p = update.Parameter; bool success;
                            switch (p.StorageType)
                            {
                                case StorageType.String: success = p.Set((string)update.After); break;
                                case StorageType.Integer: success = p.Set((int)update.After); break;
                                case StorageType.Double: success = p.Set((double)update.After); break;
                                case StorageType.ElementId: success = p.Set(RevitService.Id(Json.Integer(update.After, "value"))); break;
                                default: throw new BridgeException("unsupported_parameter", "Тип параметра не поддерживается.");
                            }
                            if (!success) throw new BridgeException("set_failed", "Revit отклонил параметр " + p.Definition.Name);
                        }
                        doc.Regenerate();
                        foreach (var update in changed)
                            if (!Equal(Raw(update.Parameter), update.After)) throw new BridgeException("verification_failed", "Revit изменил запрошенное значение; транзакция отменена.");
                        if (transaction.Commit() != TransactionStatus.Committed)
                            throw new BridgeException("transaction_rolled_back", "Revit откатил изменения: " + string.Join("; ", failures.Messages), 409);
                        // Updaters may change values at commit; keep the group open
                        // until committed values have been verified as well.
                        foreach (var update in changed)
                            if (!Equal(Raw(update.Parameter), update.After)) throw new BridgeException("verification_failed", "Значение изменилось при Commit; пакет отменён.");
                        if (group.Assimilate() != TransactionStatus.Committed) throw new BridgeException("transaction_failed", "Не удалось завершить группу транзакций.");
                    }
                    catch
                    {
                        if (transaction.GetStatus() == TransactionStatus.Started) transaction.RollBack();
                        if (group.GetStatus() == TransactionStatus.Started) group.RollBack();
                        throw;
                    }
                }
            }
            return new { dry_run = false, updates = updates.Select(u => u.Describe(Equal(u.Before, u.After) ? "unchanged" : "updated")).ToArray(), warnings = failures.Messages.ToArray(), saved_to_disk = false };
        }
        private sealed class FailureCollector : IFailuresPreprocessor
        {
            public readonly List<string> Messages = new List<string>();
            public FailureProcessingResult PreprocessFailures(FailuresAccessor accessor)
            {
                // No interactive Revit error dialog on a background request. Roll back
                // the whole batch on any failure, including warnings; preserve the cause.
                var messages = accessor.GetFailureMessages();
                foreach (var message in messages) Messages.Add(message.GetDescriptionText());
                return messages.Count > 0 ? FailureProcessingResult.ProceedWithRollBack : FailureProcessingResult.Continue;
            }
        }
    }
}
