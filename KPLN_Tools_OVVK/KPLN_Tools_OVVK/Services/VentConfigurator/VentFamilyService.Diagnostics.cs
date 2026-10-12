using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using KPLN_Tools_OVVK.Forms.Entities.VentConfigurator;

namespace KPLN_Tools_OVVK.Services.VentConfigurator
{
    internal sealed partial class VentFamilyService
    {
        private sealed class FamilyDiagnosticSnapshot
        {
            private string _typeName;
            private Dictionary<string, string> _warnings;
            private HashSet<long> _instances;
            private readonly List<string> _readErrors = new List<string>();

            internal static FamilyDiagnosticSnapshot Read(Document document)
            {
                var result = new FamilyDiagnosticSnapshot
                {
                    _typeName = document.FamilyManager.CurrentType?.Name ?? "не выбран"
                };
                try
                {
                    var warnings = new Dictionary<string, string>(StringComparer.Ordinal);
                    foreach (var message in document.GetWarnings())
                    {
                        // Уже сохранённые предупреждения фильтруем так же, как сообщения транзакции.
                        if (TransactionFailures.IsEmptyBlockDuplicate(document, message.GetFailureDefinitionId(), message.GetFailingElements(), message.GetAdditionalElements()))
                            continue;
                        string failing = Ids(message.GetFailingElements());
                        string additional = Ids(message.GetAdditionalElements());
                        string key = message.GetFailureDefinitionId().Guid.ToString() + "|" + failing + "|" + additional;
                        warnings[key] = message.GetDescriptionText() + " [ID: " + failing + "]" + (additional.Length == 0 ? "" : " [Связанные ID: " + additional + "]");
                    }

                    result._warnings = warnings;
                }
                catch (Exception ex)
                {
                    result._readErrors.Add("Предупреждения не прочитаны: " + OperationError.Message(ex));
                }

                try
                {
                    using (var collector = new FilteredElementCollector(document))
                        result._instances = new HashSet<long>(collector.OfClass(typeof(FamilyInstance)).WhereElementIsNotElementType().ToElementIds().Select(Common.IDHelper.ElIdValue));
                }
                catch (Exception ex)
                {
                    result._readErrors.Add("Экземпляры не прочитаны: " + OperationError.Message(ex));
                }

                return result;
            }

            private static string Ids(IEnumerable<ElementId> ids)
            {
                return string.Join(", ", ids.Select(Common.IDHelper.ElIdValue).Distinct().OrderBy(id => id));
            }

            internal string Report(string title)
            {
                var lines = new List<string>
                {
                    title + ":",
                    "Текущий тип: " + _typeName,
                    "Вложенных экземпляров: " + (_instances == null ? "не прочитано" : _instances.Count.ToString()),
                    "Предупреждений для проверки: " + (_warnings == null ? "не прочитано" : _warnings.Count.ToString())
                };
                if (_warnings != null)
                    lines.AddRange(_warnings.OrderBy(p => p.Key, StringComparer.Ordinal).Select(p => "  " + p.Value));
                lines.AddRange(_readErrors);
                return string.Join("\n", lines);
            }

            internal string CompareTo(FamilyDiagnosticSnapshot before)
            {
                var lines = new List<string>
                {
                    Report("После подтверждения итогового типа")
                };
                if (_instances != null && before._instances != null)
                {
                    var added = _instances.Except(before._instances).OrderBy(id => id).ToList();
                    var removed = before._instances.Except(_instances).OrderBy(id => id).ToList();
                    lines.Add("Добавленных вложенных экземпляров: " + added.Count + (added.Count == 0 ? "" : " [ID: " + string.Join(", ", added) + "]"));
                    lines.Add("Удалённых вложенных экземпляров: " + removed.Count + (removed.Count == 0 ? "" : " [ID: " + string.Join(", ", removed) + "]"));
                }

                if (_warnings != null && before._warnings != null)
                {
                    var added = _warnings.Where(p => !before._warnings.ContainsKey(p.Key)).ToList();
                    int unchanged = _warnings.Keys.Count(before._warnings.ContainsKey);
                    int removed = before._warnings.Keys.Count(key => !_warnings.ContainsKey(key));
                    lines.Add("Предупреждения: совпадают с исходными — " + unchanged + "; появились — " + added.Count + "; исчезли — " + removed + ".");
                    lines.AddRange(added.Select(p => "  Появилось: " + p.Value));
                }

                lines.Add("Сравнение относится к текущим типам до и после операции; само по себе появление предупреждения не доказывает создание экземпляра.");
                return string.Join("\n", lines);
            }
        }

        private sealed class TransactionFailures : IFailuresPreprocessor
        {
            private const string EmptyBlockFamilyName = "550_Вл_Универсальная установка_Одноуровневая_Секция_Пустой блок_(Об)";
            private readonly bool _ignoreEmptyBlockDuplicates;
            private readonly List<string> _messages = new List<string>();
            private readonly List<string> _errors = new List<string>();

            internal TransactionFailures(bool ignoreEmptyBlockDuplicates = false)
            {
                _ignoreEmptyBlockDuplicates = ignoreEmptyBlockDuplicates;
            }

            internal bool HasMessages
            {
                get
                {
                    return _messages.Count > 0;
                }
            }

            private bool IsIgnoredEmptyBlockWarning(FailuresAccessor accessor, FailureMessageAccessor message)
            {
                return _ignoreEmptyBlockDuplicates && message.GetSeverity() == FailureSeverity.Warning && IsEmptyBlockDuplicate(accessor.GetDocument(), message.GetFailureDefinitionId(), message.GetFailingElementIds(), message.GetAdditionalElementIds());
            }

            internal static bool IsEmptyBlockDuplicate(Document document, FailureDefinitionId failureId, IEnumerable<ElementId> failingIds, IEnumerable<ElementId> additionalIds)
            {
                if (document == null || !document.IsFamilyDocument || failureId.Guid != BuiltInFailures.OverlapFailures.DuplicateInstances.Guid)
                    return false;
                var failing = failingIds.Distinct().ToList();
                if (failing.Count < 2)
                    return false;
                // Проверяем реальные элементы, а не локализованный текст или ID из одного отчёта.
                return failing.Concat(additionalIds).Distinct().All(id =>
                {
                    var symbol = (document.GetElement(id) as FamilyInstance)?.Symbol;
                    return symbol != null && string.Equals(symbol.FamilyName, EmptyBlockFamilyName, StringComparison.OrdinalIgnoreCase) && string.Equals(symbol.Name, "Пустой блок", StringComparison.OrdinalIgnoreCase);
                });
            }

            public FailureProcessingResult PreprocessFailures(FailuresAccessor accessor)
            {
                bool error = false;
                foreach (var message in accessor.GetFailureMessages().ToList())
                {
                    if (IsIgnoredEmptyBlockWarning(accessor, message))
                    {
                        // Полностью игнорируем: не показываем окно и не добавляем сообщение в подробности.
                        accessor.DeleteWarning(message);
                        continue;
                    }

                    var severity = message.GetSeverity();
                    string description = "[" + severity + "; " + message.GetFailureDefinitionId().Guid + "] " + message.GetDescriptionText();
                    var ids = message.GetFailingElementIds().Concat(message.GetAdditionalElementIds()).Select(Common.IDHelper.ElIdValue).Distinct().ToList();
                    if (ids.Count > 0)
                        description += " [ID: " + string.Join(", ", ids) + "]";
                    if (!string.IsNullOrWhiteSpace(description) && !_messages.Contains(description))
                        _messages.Add(description);
                    if (severity == FailureSeverity.Error || severity == FailureSeverity.DocumentCorruption)
                    {
                        error = true;
                        if (!_errors.Contains(description))
                            _errors.Add(description);
                    }
                }

                // Ошибочный результат не сохраняем и не предлагаем снимать зависимости ради продолжения.
                // Остальные предупреждения остаются в обработке Revit; в окно ошибки попадают только ошибки.
                return error ? FailureProcessingResult.ProceedWithRollBack : FailureProcessingResult.Continue;
            }

            internal string Describe(string fallback)
            {
                return _messages.Count == 0 ? fallback : fallback + "\n" + string.Join("\n", _messages);
            }

            internal string DescribeErrors(string fallback)
            {
                return _errors.Count == 0 ? fallback : fallback + "\n" + string.Join("\n", _errors);
            }
        }

        private static string LoadType(Document target, string path, string typeName)
        {
            var loadOptions = new FamilyLoadOptions(Path.GetFileNameWithoutExtension(path));
            using (var transaction = new Transaction(target, "Загрузить тип вентиляционной установки"))
            {
                if (transaction.Start() != TransactionStatus.Started)
                    throw new InvalidOperationException("Не удалось начать загрузку в проект.");
                var reportedFailures = new TransactionFailures();
                var failures = transaction.GetFailureHandlingOptions();
                failures.SetFailuresPreprocessor(reportedFailures);
                failures.SetClearAfterRollback(true);
                failures.SetForcedModalHandling(true);
                transaction.SetFailureHandlingOptions(failures);
                FamilySymbol symbol;
                bool loaded = target.LoadFamilySymbol(path, typeName, loadOptions, out symbol);
                if (loadOptions.Cancelled)
                {
                    transaction.RollBack();
                    return "Загрузка в проект отменена.";
                }
                if (!loaded || symbol == null)
                {
                    transaction.RollBack();
                    return "Revit не выполнил загрузку типа в проект. Возможно, такая версия типа уже загружена.";
                }
                if (!symbol.IsActive)
                    symbol.Activate();
                if (transaction.Commit() != TransactionStatus.Committed)
                    throw new InvalidOperationException(reportedFailures.DescribeErrors("Revit отменил загрузку типа в проект."));
            }
            return "Тип «" + typeName + "» загружен в проект «" + target.Title + "». Экземпляр не размещён.";
        }
    }
}
