using Autodesk.Revit.DB;
using KPLN_Library_ExtensibleStorage;
using KPLN_OpeningHoleManager.Core;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace KPLN_OpeningHoleManager.Services
{
    internal static class OpeningHistoryService
    {
        internal const string CacheTransactionName = "KPLN: Кэш истории отверстий";
        private static readonly string[] Fields =
        {
            "KPLN_OHM_CreatedDate", "KPLN_OHM_CreatedAuthor", "KPLN_OHM_CreatorLogin",
            "KPLN_OHM_ElementUniqueId", "KPLN_OHM_ConfirmedFingerprint", "KPLN_OHM_HistoryXml"
        };
        private static readonly ExtensibleStorageBuilder Storage = new ExtensibleStorageBuilder(
            new Guid("0060D224-F58A-40AF-A206-9B6C246C3877"), Fields, "KPLN_OHM_OpeningHistory_v2");

        private static readonly string[] FamilyPrefixes = ARKRElemsWorker.OpeningNames_StartWith;
        internal static readonly ElementFilter CandidateFilter = new LogicalAndFilter(
            new ElementClassFilter(typeof(FamilyInstance)), ARKRElemsWorker.OpeningElemCatLogicalOrFilter);


        internal static bool IsValidId(ElementId id) => id != null && !id.Equals(ElementId.InvalidElementId);

        internal static bool IsOpeningFamily(string name) => !string.IsNullOrEmpty(name)
            && FamilyPrefixes.Any(prefix => name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));

        internal static bool IsOpening(Element element)
        {
            FamilyInstance instance = element as FamilyInstance;
            return instance != null && instance.IsValidObject
                && ARKRElemsWorker.OpeningElemCatLogicalOrFilter.PassesFilter(instance)
                && IsOpeningFamily(instance.Symbol?.FamilyName);
        }

        internal static OpeningHistoryRow[] Collect(Document doc, bool cacheMissing, ISet<string> cacheElementIds = null,
            IReadOnlyDictionary<string, OpeningCreationInfo> observedCreations = null, IEnumerable<ElementId> elementIds = null)
        {
            // Полный сбор нужен при открытии/ручном обновлении. Для событий читаем только переданные ID.
            IEnumerable<Element> candidates = elementIds == null
                ? new FilteredElementCollector(doc).WherePasses(CandidateFilter).WhereElementIsNotElementType()
                : elementIds.Select(doc.GetElement);
            FamilyInstance[] openings = candidates.Where(IsOpening).Cast<FamilyInstance>()
                .OrderBy(e => e.Id.ToString().Length).ThenBy(e => e.Id.ToString(), StringComparer.Ordinal).ToArray();


            var records = new Dictionary<string, OpeningHistoryRecord>();
            var notes = new Dictionary<string, string>();
            var missing = new List<FamilyInstance>();
            foreach (FamilyInstance element in openings)
            {
                try
                {
                    bool changed;
                    OpeningCreationInfo creation = null;
                    observedCreations?.TryGetValue(element.UniqueId, out creation);
                    records[element.UniqueId] = Read(element, out changed, creation);
                    if (changed)
                        missing.Add(element);
                }
                catch (Exception ex)
                {
                    notes[element.UniqueId] = "История не прочитана: " + OpeningUserMessage.FromException(ex);
                }
            }


            // Автоматически записываем только новые отверстия из события; ручное обновление кэширует все.
            FamilyInstance[] toWrite = missing.Where(element => cacheMissing
                || cacheElementIds != null && cacheElementIds.Contains(element.Id.ToString())).ToArray();
            foreach (FamilyInstance element in missing.Except(toWrite))
                notes[element.UniqueId] = doc.IsReadOnly ? "Модель доступна только для чтения" : "Данные ещё не записаны — нажмите «Обновить»";

            if (toWrite.Length > 0 && !doc.IsReadOnly && !doc.IsModifiable)
            {
                using (Transaction tx = new Transaction(doc, CacheTransactionName))
                {
                    var failures = new CacheFailures();
                    var pending = new List<FamilyInstance>();
                    try
                    {
                        tx.Start();
                        tx.SetFailureHandlingOptions(tx.GetFailureHandlingOptions()
                            .SetFailuresPreprocessor(failures).SetClearAfterRollback(true));
                        foreach (FamilyInstance element in toWrite)
                        {
                            string id = element.UniqueId;
                            if (doc.IsWorkshared && WorksharingUtils.GetCheckoutStatus(doc, element.Id) == CheckoutStatus.OwnedByOtherUser)
                            {
                                notes[id] = "Не записано: элемент занят другим пользователем";
                                continue;
                            }
                            using (SubTransaction sub = new SubTransaction(doc))
                            {
                                try
                                {
                                    sub.Start();
                                    Write(element, records[id]);
                                    if (sub.Commit() == TransactionStatus.Committed)
                                        pending.Add(element);
                                    else
                                        notes[id] = "Не удалось записать историю";
                                }
                                catch (Exception ex)
                                {
                                    if (sub.GetStatus() == TransactionStatus.Started) sub.RollBack();
                                    notes[id] = "Не записано: " + KPLN_OpeningHoleManager.Services.OpeningUserMessage.FromException(ex);
                                }
                            }
                        }
                        if (tx.Commit() != TransactionStatus.Committed)
                            foreach (FamilyInstance element in pending)
                                notes[element.UniqueId] = "Кэш не записан: " + failures.ErrorMessage;
                    }
                    catch (Exception ex)
                    {
                        if (tx.GetStatus() == TransactionStatus.Started) tx.RollBack();
                        foreach (FamilyInstance element in toWrite)
                        {
                            // Сохраняем точную причину для занятых элементов и ошибок отдельной записи.
                            if (!notes.ContainsKey(element.UniqueId))
                                notes[element.UniqueId] = "Кэш не записан: " + KPLN_OpeningHoleManager.Services.OpeningUserMessage.FromException(ex);
                        }
                    }
                }
            }
            else
                foreach (FamilyInstance element in toWrite)
                    notes[element.UniqueId] = doc.IsReadOnly ? "Модель доступна только для чтения" : "Данные ещё не записаны — нажмите «Обновить»";



            // Данные карточек и зависимости для последующих точечных обновлений.
            return openings.Select(element =>
            {
                OpeningHistoryRecord record;
                records.TryGetValue(element.UniqueId, out record);
                string note;
                notes.TryGetValue(element.UniqueId, out note);
                Level level = doc.GetElement(element.LevelId) as Level;
                if (level == null && element.Host != null)
                    level = doc.GetElement(element.Host.LevelId) as Level;
                return new OpeningHistoryRow
                {
                    RevitId = element.Id,
                    DependencyIds = GetDependencyIds(element),
                    UniqueId = element.UniqueId,
                    ElementId = element.Id.ToString(),
                    TypeName = element.Symbol.FamilyName + ": " + element.Name,
                    LevelId = level?.Id.ToString() ?? string.Empty,
                    LevelName = level?.Name ?? "Без уровня",
                    LevelElevation = level?.Elevation ?? double.MinValue,
                    CreatedDate = record?.CreatedDate ?? OpeningHistoryRecord.UnknownDate,
                    CreatedAuthor = record?.CreatedAuthor ?? OpeningHistoryRecord.UnknownAuthor,
                    Status = string.IsNullOrEmpty(record?.ConfirmedFingerprint) ? "Ожидает подтверждения"
                        : record.ConfirmedFingerprint == Fingerprint(element) ? "Подтверждено" : "Изменено после подтверждения",
                    CacheNote = note ?? string.Empty,
                    Entries = record?.Entries.AsEnumerable().Reverse().ToArray() ?? new OpeningHistoryEntry[0],
                };
            }).ToArray();
        }

        private static string[] GetDependencyIds(FamilyInstance opening)
        {
            Element host = opening.Host;
            var ids = new[]
            {
                opening.GetTypeId(), opening.Symbol?.Family?.Id, opening.LevelId,
                host?.Id, host?.GetTypeId(), host?.LevelId,
                (host as FamilyInstance)?.Symbol?.Family?.Id,
            };

            return ids.Where(IsValidId).Select(id => id.ToString()).Distinct().ToArray();
        }


        internal static OpeningCreationInfo CaptureCreation(Element element)
        {
            string login = Creator(element);
            return new OpeningCreationInfo
            {
                Date = Now(),
                CreatorLogin = string.IsNullOrWhiteSpace(login) ? element.Document.Application.Username : login,
            };
        }

        private static OpeningHistoryRecord Read(Element element, out bool changed, OpeningCreationInfo creation = null)
        {
            string[] values = Fields.Select(field => Storage.GetStorageData_TextData(field, element)).ToArray();
            // Копирование в Revit переносит Extensible Storage: чужое подтверждение не наследуем.
            bool copied = !string.IsNullOrEmpty(values[3]) && values[3] != element.UniqueId;
            OpeningHistoryRecord record = copied ? new OpeningHistoryRecord() : new OpeningHistoryRecord
            {
                CreatedDate = values[0], CreatedAuthor = values[1], CreatorLogin = values[2],
                ElementUniqueId = values[3], ConfirmedFingerprint = values[4], Entries = OpeningHistoryRecord.Parse(values[5]),
            };
            // Дата получена при появлении элемента, а не при последующем выполнении очереди.
            // Собственную историю, уже записанную командой плагина, не заменяем общим событием.
            bool observed = creation != null && (copied || string.IsNullOrWhiteSpace(record.ElementUniqueId));
            if (observed)
                record = new OpeningHistoryRecord
                {
                    CreatedDate = creation.Date,
                    CreatorLogin = creation.CreatorLogin,
                    CreatedAuthor = OpeningAuthorResolver.Resolve(creation.CreatorLogin),
                };

            bool compacted = NormalizeCreationHistory(record);
            changed = observed || copied || compacted || string.IsNullOrWhiteSpace(record.CreatedDate) || string.IsNullOrWhiteSpace(record.CreatedAuthor)
                || string.IsNullOrWhiteSpace(record.ElementUniqueId) || record.Entries.Count == 0;
            if (string.IsNullOrWhiteSpace(record.CreatedDate)) record.CreatedDate = OpeningHistoryRecord.UnknownDate;
            if (string.IsNullOrWhiteSpace(record.CreatedAuthor))
            {
                if (string.IsNullOrWhiteSpace(record.CreatorLogin)) record.CreatorLogin = Creator(element);
                record.CreatedAuthor = OpeningAuthorResolver.Resolve(record.CreatorLogin);
            }
            record.ElementUniqueId = element.UniqueId;
            if (record.Entries.Count == 0)
                record.Entries.Add(new OpeningHistoryEntry
                {
                    Date = record.CreatedDate, Author = record.CreatedAuthor,
                    Action = "Создано", Details = string.Empty,
                });
            return record;
        }

        private static string Creator(Element element)
        {
            if (!element.Document.IsWorkshared) return string.Empty;
            try { return WorksharingUtils.GetWorksharingTooltipInfo(element.Document, element.Id).Creator; }
            catch { return string.Empty; }
        }

        private static void Write(Element element, OpeningHistoryRecord record)
        {
            Storage.SetStorageData_TextData(element, new Dictionary<string, string>
            {
                { Fields[0], record.CreatedDate }, { Fields[1], record.CreatedAuthor }, { Fields[2], record.CreatorLogin },
                { Fields[3], element.UniqueId }, { Fields[4], record.ConfirmedFingerprint }, { Fields[5], record.ToXml() },
            });

        }

        private static bool NormalizeCreationHistory(OpeningHistoryRecord record)
        {
            // Старые объединения переносили всю историю временных исходных элементов.
            // Оставляем события самого итогового отверстия, включая его подтверждения.
            bool changed = record.Entries.RemoveAll(entry =>
                (entry.Details ?? string.Empty).StartsWith("Исходное отверстие ID ", StringComparison.Ordinal)) > 0;
            foreach (OpeningHistoryEntry entry in record.Entries)
            {
                if (entry.Action == "Создание")
                {
                    entry.Action = "Создано";
                    if (entry.Details == "Скопированный элемент; дата создания неизвестна" || entry.Details == "Данные существующего отверстия")
                        entry.Details = string.Empty;
                    changed = true;
                }
                else if (entry.Action == "Создано при объединении отверстий")
                {
                    entry.Action = "Создано плагином";
                    // Точная команда в прежнем формате не сохранялась.
                    entry.Details = "Сценарий: Создание с объединением";
                    changed = true;
                }
                else if (entry.Action == "Создано плагином"
                    && (string.IsNullOrWhiteSpace(entry.Details) || entry.Details.StartsWith("ID ", StringComparison.Ordinal)))
                {
                    entry.Details = "Сценарий: Не определён";
                    changed = true;
                }
            }
            return changed;
        }

        internal static void RecordCreated(FamilyInstance instance, string scenario)
        {
            string login = instance.Document.Application.Username;
            var record = new OpeningHistoryRecord
            {
                CreatedDate = Now(), CreatedAuthor = OpeningAuthorResolver.Resolve(login), CreatorLogin = login,
                ElementUniqueId = instance.UniqueId,
            };
            record.Entries.Add(new OpeningHistoryEntry
            {
                Date = record.CreatedDate, Author = record.CreatedAuthor,
                Action = "Создано плагином",
                Details = "Сценарий: " + (string.IsNullOrWhiteSpace(scenario) ? "Не определён" : scenario),
            });
            Write(instance, record);
        }

        internal static void RecordChange(Element element, string action)
        {
            bool ignored;
            OpeningHistoryRecord record = Read(element, out ignored);
            record.ConfirmedFingerprint = string.Empty;
            record.Entries.Add(new OpeningHistoryEntry { Date = Now(), Author = OpeningAuthorResolver.Resolve(element.Document.Application.Username), Action = action });
            Write(element, record);
        }

        internal static bool IsConfirmed(Element element)
        {
            bool ignored;
            OpeningHistoryRecord record = Read(element, out ignored);
            return record.ConfirmedFingerprint == Fingerprint(element);
        }

        internal static bool Confirm(Element element, string bimAuthor = null)
        {
            bool ignored;
            OpeningHistoryRecord record = Read(element, out ignored);
            string fingerprint = Fingerprint(element);
            if (record.ConfirmedFingerprint == fingerprint) return false;
            record.ConfirmedFingerprint = fingerprint;
            record.Entries.Add(new OpeningHistoryEntry
            {
                Date = Now(),
                Author = bimAuthor ?? OpeningAuthorResolver.Resolve(element.Document.Application.Username),
                Action = bimAuthor == null ? "Подтверждено пользователем" : "Подтверждено BIM-отделом",
                Details = bimAuthor == null ? string.Empty : "Массовое подтверждение",
            });
            Write(element, record);
            return true;
        }

        internal static string Fingerprint(Element element)
        {
            var text = new StringBuilder(element.GetTypeId().ToString());
            FamilyInstance instance = element as FamilyInstance;
            text.Append('|').Append(instance?.Host?.UniqueId).Append('|').Append(element.LevelId);
            LocationPoint point = element.Location as LocationPoint;
            if (point != null)
                foreach (double value in new[] { point.Point.X, point.Point.Y, point.Point.Z, point.Rotation })
                    text.Append('|').Append(value.ToString("0.#########", CultureInfo.InvariantCulture));
            if (instance != null) text.Append('|').Append(instance.FacingFlipped).Append('|').Append(instance.HandFlipped);
            BoundingBoxXYZ bounds = element.get_BoundingBox(null);
            if (bounds != null)
                foreach (double value in new[] { bounds.Min.X, bounds.Min.Y, bounds.Min.Z, bounds.Max.X, bounds.Max.Y, bounds.Max.Z })
                    text.Append('|').Append(value.ToString("0.#########", CultureInfo.InvariantCulture));
            AppendParameters(text, element.Parameters.Cast<Parameter>());
            if (instance != null)
                AppendParameters(text, instance.Symbol.Parameters.Cast<Parameter>());
            using (SHA256 sha = SHA256.Create()) return Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(text.ToString())));
        }

        private static void AppendParameters(StringBuilder text, IEnumerable<Parameter> parameters)
        {
            text.Append("|parameters:");
            foreach (Parameter parameter in parameters.OrderBy(p => p.Id.ToString(), StringComparer.Ordinal))
            {
                if (parameter.StorageType == StorageType.Double)
                    text.Append('|').Append(parameter.Id).Append('=').Append(parameter.AsDouble().ToString("0.#########", CultureInfo.InvariantCulture));
                else if (parameter.StorageType == StorageType.Integer && !parameter.IsReadOnly)
                    text.Append('|').Append(parameter.Id).Append('=').Append(parameter.AsInteger());
            }
        }

        private static string Now() => DateTimeOffset.Now.ToString("dd.MM.yyyy HH:mm:ss zzz", CultureInfo.InvariantCulture);

        internal sealed class CacheFailures : IFailuresPreprocessor
        {
            private readonly List<string> _errors = new List<string>();

            internal string ErrorMessage => _errors.Count == 0
                ? "Revit отменил транзакцию без описания причины"
                : string.Join(Environment.NewLine, _errors.Distinct());

            public FailureProcessingResult PreprocessFailures(FailuresAccessor accessor)
            {
                foreach (FailureMessageAccessor failure in accessor.GetFailureMessages())
                {
                    FailureSeverity severity = failure.GetSeverity();
                    if (severity == FailureSeverity.Warning)
                    {
                        // Записываем только метаданные: предупреждение не требует отката или изменения геометрии.
                        accessor.DeleteWarning(failure);
                    }
                    else if (severity != FailureSeverity.None)
                    {
                        string description = OpeningUserMessage.FromFailure(failure.GetDescriptionText());
                        string ids = string.Join(", ", failure.GetFailingElementIds().Select(id => id.ToString()));
                        _errors.Add(string.IsNullOrWhiteSpace(ids) ? description : description + " (ID: " + ids + ")");
                    }
                }

                // Ошибки не разрешаем удалением/отсоединением элементов: сохраняем модель и объясняем отказ.
                return _errors.Count > 0
                    ? FailureProcessingResult.ProceedWithRollBack
                    : FailureProcessingResult.Continue;
            }
        }
    }
}
