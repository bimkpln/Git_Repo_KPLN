using Autodesk.Revit.DB;
using KPLN_Library_ExtensibleStorage;
using KPLN_OpeningHoleManager.Core;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace KPLN_OpeningHoleManager.Services
{
    internal static class OpeningHistoryService
    {
        internal const string CacheTransactionName = "KPLN: Снимки и история отверстий";
        private const string Prefix = "KPLN_OHM_";
        private static readonly string[] SharedFields = { "ElementUniqueId", "CreatedAuthor", "CreatorLogin", "HistoryXml" };
        private static readonly string[] AuditFields =
        {
            "Status", "StatusAuthor", "StatusLogin", "ConfirmedDate", "ConfirmedAuthor", "ConfirmedLogin",
        };
        private static readonly string[] Fields = SharedFields.Select(field => Prefix + field)
            .Concat(OpeningSnapshot.Fields.Select(field => Prefix + "G1_" + field))
            .Concat(OpeningSnapshot.Fields.Select(field => Prefix + "G2_" + field)).ToArray();
        private static readonly ExtensibleStorageBuilder Storage = new ExtensibleStorageBuilder(
            new Guid("F940943D-76DB-44F9-B65E-C4A75A3ACCE8"), Fields, "KPLN_OHM_OpeningHistory_v3");

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


        // Обновление поколений вызывается только при открытии панели и ручном обновлении.
        // После подтверждения перечитываем указанные карточки без нового снимка и записи.
        internal static OpeningHistoryRow[] Collect(Document doc, bool updateSnapshots, IEnumerable<ElementId> elementIds = null)
        {
            IEnumerable<Element> candidates = elementIds == null
                ? new FilteredElementCollector(doc).WherePasses(CandidateFilter).WhereElementIsNotElementType()
                : elementIds.Select(doc.GetElement);
            FamilyInstance[] openings = candidates.Where(IsOpening).Cast<FamilyInstance>()
                .OrderBy(e => e.Id.ToString().Length).ThenBy(e => e.Id.ToString(), StringComparer.Ordinal).ToArray();

            var records = new Dictionary<string, OpeningHistoryRecord>();
            var notes = new Dictionary<string, string>();
            var toWrite = new List<FamilyInstance>();

            foreach (FamilyInstance opening in openings)
            {
                try
                {
                    OpeningHistoryRecord record = Read(opening);
                    Dictionary<string, string> before = Serialize(record);

                    if (updateSnapshots)
                        UpdateSnapshot(opening, record);

                    records[opening.UniqueId] = record;
                    if (updateSnapshots && !SameValues(before, Serialize(record)))
                        toWrite.Add(opening);
                }
                catch (Exception ex)
                {
                    notes[opening.UniqueId] = "Данные отверстия не прочитаны: " + OpeningUserMessage.FromException(ex);
                }
            }

            SaveSnapshots(doc, toWrite, records, notes);

            return openings.Select(opening => CreateRow(opening,
                records.TryGetValue(opening.UniqueId, out OpeningHistoryRecord record) ? record : null,
                notes.TryGetValue(opening.UniqueId, out string note) ? note : string.Empty)).ToArray();
        }

        private static void SaveSnapshots(Document doc, List<FamilyInstance> openings,
            Dictionary<string, OpeningHistoryRecord> records, Dictionary<string, string> notes)
        {
            if (openings.Count == 0)
                return;

            if (doc.IsReadOnly || doc.IsModifiable)
            {
                foreach (FamilyInstance opening in openings)
                    notes[opening.UniqueId] = doc.IsReadOnly ? "Модель доступна только для чтения"
                        : "Снимок не записан: завершите текущую операцию и нажмите «Обновить»";
                return;
            }


            using (Transaction tx = new Transaction(doc, CacheTransactionName))
            {
                var failures = new CacheFailures();
                var pending = new List<FamilyInstance>();
                try
                {
                    tx.Start();
                    tx.SetFailureHandlingOptions(tx.GetFailureHandlingOptions()
                        .SetFailuresPreprocessor(failures).SetClearAfterRollback(true));

                    foreach (FamilyInstance opening in openings)
                    {
                        string id = opening.UniqueId;
                        if (doc.IsWorkshared && WorksharingUtils.GetCheckoutStatus(doc, opening.Id) == CheckoutStatus.OwnedByOtherUser)
                        {
                            notes[id] = "Не записано: элемент занят другим пользователем";
                            continue;
                        }

                        using (SubTransaction sub = new SubTransaction(doc))
                        {
                            try
                            {
                                sub.Start();
                                Write(opening, records[id]);
                                if (sub.Commit() == TransactionStatus.Committed)
                                    pending.Add(opening);
                                else
                                    notes[id] = "Не удалось записать снимок отверстия";
                            }
                            catch (Exception ex)
                            {
                                if (sub.GetStatus() == TransactionStatus.Started)
                                    sub.RollBack();
                                notes[id] = "Не записано: " + OpeningUserMessage.FromException(ex);
                            }
                        }
                    }

                    if (pending.Count == 0)
                        tx.RollBack();
                    else if (tx.Commit() != TransactionStatus.Committed)
                        foreach (FamilyInstance opening in pending)
                            notes[opening.UniqueId] = "Снимок не записан: " + failures.ErrorMessage;
                }
                catch (Exception ex)
                {
                    if (tx.GetStatus() == TransactionStatus.Started)
                        tx.RollBack();

                    foreach (FamilyInstance opening in openings)
                        if (!notes.ContainsKey(opening.UniqueId))
                            notes[opening.UniqueId] = "Снимок не записан: " + OpeningUserMessage.FromException(ex);
                }
            }
        }


        private static OpeningHistoryRow CreateRow(FamilyInstance opening, OpeningHistoryRecord record, string note)
        {
            Level level = opening.Document.GetElement(opening.LevelId) as Level;
            if (level == null && opening.Host != null)
                level = opening.Document.GetElement(opening.Host.LevelId) as Level;

            OpeningSnapshot snapshot = record?.Current;
            string statusAuthor = snapshot?["StatusAuthor"] ?? string.Empty;
            if (snapshot?["Status"] == OpeningStatuses.New && string.IsNullOrWhiteSpace(statusAuthor))
                statusAuthor = OpeningHistoryRecord.UnknownAuthor;

            return new OpeningHistoryRow
            {
                RevitId = opening.Id,
                UniqueId = opening.UniqueId,
                ElementId = opening.Id.ToString(),
                TypeName = opening.Symbol.FamilyName + ": " + opening.Name,
                LevelId = level?.Id.ToString() ?? string.Empty,
                LevelName = level?.Name ?? "Без уровня",
                LevelElevation = level?.Elevation ?? double.MinValue,
                GeometryDetails = OpeningSnapshotService.DescribeGeometry(snapshot),
                CreatedAuthor = string.IsNullOrWhiteSpace(record?.CreatedAuthor) ? OpeningHistoryRecord.UnknownAuthor : record.CreatedAuthor,
                StatusCode = snapshot?["Status"] ?? OpeningStatuses.New,
                StatusAuthor = statusAuthor,
                ConfirmedDate = snapshot?["ConfirmedDate"] ?? string.Empty,
                ConfirmedAuthor = snapshot?["ConfirmedAuthor"] ?? string.Empty,
                CacheNote = note,
                Entries = record?.DisplayChanges() ?? new OpeningHistoryEntry[0],
            };
        }

        internal static OpeningHistoryRecord Read(Element element)
        {
            // Читаем только новую схему. Entity скопированного элемента относится к другому UniqueId.
            var values = Storage.GetStorageData_TextData(element, Fields)
                .ToDictionary(pair => pair.Key, pair => pair.Value ?? string.Empty);
            if (values[Prefix + "ElementUniqueId"] != element.UniqueId)
                return new OpeningHistoryRecord();

            var record = new OpeningHistoryRecord
            {
                ElementUniqueId = values[Prefix + "ElementUniqueId"],
                CreatedAuthor = values[Prefix + "CreatedAuthor"],
                CreatorLogin = values[Prefix + "CreatorLogin"],
                Current = ReadSnapshot(values, "G1_"),
                Previous = ReadSnapshot(values, "G2_"),
            };
            record.ReadHistory(values[Prefix + "HistoryXml"]);

            return record;
        }

        private static OpeningSnapshot ReadSnapshot(Dictionary<string, string> values, string generation)
        {
            if (string.IsNullOrEmpty(values[Prefix + generation + "Status"]))
                return null;

            var snapshot = new OpeningSnapshot();
            foreach (string field in OpeningSnapshot.Fields)
                snapshot[field] = values[Prefix + generation + field];

            return snapshot;
        }

        private static Dictionary<string, string> Serialize(OpeningHistoryRecord record)
        {
            var values = new Dictionary<string, string>
            {
                { Prefix + "ElementUniqueId", record.ElementUniqueId },
                { Prefix + "CreatedAuthor", record.CreatedAuthor },
                { Prefix + "CreatorLogin", record.CreatorLogin },
                { Prefix + "HistoryXml", record.HistoryXml() },
            };
            foreach (string field in OpeningSnapshot.Fields)
            {
                values[Prefix + "G1_" + field] = record.Current?[field] ?? string.Empty;
                values[Prefix + "G2_" + field] = record.Previous?[field] ?? string.Empty;
            }

            return values;
        }

        private static bool SameValues(Dictionary<string, string> left, Dictionary<string, string> right) =>
            left.Count == right.Count && left.All(pair => right.TryGetValue(pair.Key, out string value) && pair.Value == value);

        private static void Write(Element element, OpeningHistoryRecord record) =>
            Storage.SetStorageData_TextData(element, Serialize(record));


        private static string Creator(Element element)
        {
            if (!element.Document.IsWorkshared)
                return string.Empty;
            try { return WorksharingUtils.GetWorksharingTooltipInfo(element.Document, element.Id).Creator; }
            catch { return string.Empty; }
        }

        private static void FillCreator(Element element, OpeningHistoryRecord record)
        {
            // Не подменяем создателя тем, кто открыл окно или подтвердил отверстие.
            if (!string.IsNullOrWhiteSpace(record.CreatedAuthor))
                return;

            if (string.IsNullOrWhiteSpace(record.CreatorLogin))
                record.CreatorLogin = Creator(element);

            if (!string.IsNullOrWhiteSpace(record.CreatorLogin))
                record.CreatedAuthor = OpeningAuthorResolver.Resolve(record.CreatorLogin);
        }

        private static void SetStatus(OpeningSnapshot snapshot, string status, string login, string author)
        {
            snapshot["Status"] = status;
            snapshot["StatusLogin"] = login;
            snapshot["StatusAuthor"] = author;
        }

        private static void SetNewStatus(OpeningSnapshot snapshot, OpeningHistoryRecord record)
        {
            // Статус «Новое» принадлежит создателю отверстия, а не пользователю,
            // который впервые открыл менеджер или обновил список.
            SetStatus(snapshot, OpeningStatuses.New, record.CreatorLogin, record.CreatedAuthor);
        }

        private static void UpdateSnapshot(FamilyInstance opening, OpeningHistoryRecord record)
        {
            FillCreator(opening, record);
            record.ElementUniqueId = opening.UniqueId;
            OpeningSnapshot current = OpeningSnapshotService.Capture(opening);
            string login = opening.Document.Application.Username;
            string author = OpeningAuthorResolver.Resolve(login);

            if (record.Current == null)
            {
                record.Current = current;
                SetNewStatus(current, record);
                record.Append(new OpeningHistoryEntry
                {
                    Action = "Новое отверстие",
                    Author = record.CreatedAuthor,
                    Details = OpeningSnapshotService.Describe(current),
                });
                return;
            }


            // Исправляем прежнее авторство «Нового» до переноса в G2.
            // Это уточнение метаданных не является изменением самого отверстия.
            if (record.Current["Status"] == OpeningStatuses.New)
                SetNewStatus(record.Current, record);

            record.Previous = record.Current.Copy();
            foreach (string field in AuditFields)
                current[field] = record.Previous[field];
            record.Current = current;

            string[] changes = OpeningSnapshotService.Compare(record.Previous, current);
            if (changes.Length == 0)
                return;

            SetStatus(current, OpeningStatuses.Changed, login, author);
            string transition = record.Previous["Status"] == OpeningStatuses.Changed ? string.Empty
                : "Статус: " + OpeningStatuses.Get(record.Previous["Status"]).Name + " → Изменено" + Environment.NewLine;
            record.Append(new OpeningHistoryEntry
            {
                Action = OpeningHistoryEntry.ChangesDetected,
                Author = author,
                Details = transition + string.Join(Environment.NewLine, changes),
            });
        }


        internal static void RecordCreated(FamilyInstance instance, string scenario)
        {
            string login = instance.Document.Application.Username;
            string author = OpeningAuthorResolver.Resolve(login);
            var record = new OpeningHistoryRecord
            {
                ElementUniqueId = instance.UniqueId,
                CreatorLogin = login,
                CreatedAuthor = author,
                Current = OpeningSnapshotService.Capture(instance),
            };
            SetNewStatus(record.Current, record);
            record.Append(new OpeningHistoryEntry
            {
                Action = "Создано плагином",
                Author = author,
                Details = "Сценарий: " + (string.IsNullOrWhiteSpace(scenario) ? "Не определён" : scenario)
                    + Environment.NewLine + OpeningSnapshotService.Describe(record.Current),
            });
            Write(instance, record);
        }

        internal static void RecordChange(Element element, string action)
        {
            OpeningHistoryRecord record = Read(element);
            if (record.Current == null)
                UpdateSnapshot((FamilyInstance)element, record);

            // Не сдвигаем поколения во время команды: изменения геометрии обнаружит ручное обновление.
            record.Append(new OpeningHistoryEntry
            {
                Action = action,
                Author = OpeningAuthorResolver.Resolve(element.Document.Application.Username),
            });
            Write(element, record);
        }

        internal static bool IsConfirmed(Element element)
        {
            OpeningHistoryRecord record = Read(element);
            return record.Current?["Status"] == OpeningStatuses.Confirmed
                && OpeningSnapshotService.Compare(record.Current, OpeningSnapshotService.Capture((FamilyInstance)element)).Length == 0;
        }

        internal static bool Confirm(Element element, string bimAuthor = null)
        {
            OpeningHistoryRecord record = Read(element);
            OpeningSnapshot actual = OpeningSnapshotService.Capture((FamilyInstance)element);
            if (record.Current?["Status"] == OpeningStatuses.Confirmed
                && OpeningSnapshotService.Compare(record.Current, actual).Length == 0)
                return false;

            // Проверяем геометрию при самом подтверждении, даже если карточка ещё не обновлена.
            UpdateSnapshot((FamilyInstance)element, record);
            string login = element.Document.Application.Username;
            string author = bimAuthor ?? OpeningAuthorResolver.Resolve(login);
            string previousStatus = OpeningStatuses.Get(record.Current["Status"]).Name;
            SetStatus(record.Current, OpeningStatuses.Confirmed, login, author);
            record.Current["ConfirmedDate"] = DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss", CultureInfo.InvariantCulture);
            record.Current["ConfirmedAuthor"] = author;
            record.Current["ConfirmedLogin"] = login;
            record.Append(new OpeningHistoryEntry
            {
                Date = record.Current["ConfirmedDate"],
                Author = author,
                Action = bimAuthor == null ? OpeningHistoryEntry.ConfirmedByUser : OpeningHistoryEntry.ConfirmedByBim,
                Details = "Статус: " + previousStatus + " → Подтверждено",
            });
            Write(element, record);
            return true;
        }


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
                        accessor.DeleteWarning(failure);
                    else if (severity != FailureSeverity.None)
                    {
                        string description = OpeningUserMessage.FromFailure(failure.GetDescriptionText());
                        string ids = string.Join(", ", failure.GetFailingElementIds().Select(id => id.ToString()));
                        _errors.Add(string.IsNullOrWhiteSpace(ids) ? description : description + " (ID: " + ids + ")");
                    }
                }

                // Ошибки не разрешаем удалением/отсоединением элементов: откатываем метаданные целиком.
                return _errors.Count > 0 ? FailureProcessingResult.ProceedWithRollBack : FailureProcessingResult.Continue;
            }
        }
    }
}
