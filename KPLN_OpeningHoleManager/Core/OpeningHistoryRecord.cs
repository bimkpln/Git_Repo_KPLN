using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace KPLN_OpeningHoleManager.Core
{
    public static class OpeningStatuses
    {
        public const string New = "new";
        public const string Changed = "changed";
        public const string Confirmed = "confirmed";

        // Код хранится в модели; название и цвет меняются независимо от схемы хранилища.
        public static IReadOnlyList<OpeningStatusDefinition> All { get; } = new[]
        {
            new OpeningStatusDefinition(New, "Новое", "#FF6B6B"),
            new OpeningStatusDefinition(Changed, "Изменено", "#FFAD5C"),
            new OpeningStatusDefinition(Confirmed, "Подтверждено", "#81C995"),
        };

        public static OpeningStatusDefinition Get(string code) => All.FirstOrDefault(item => item.Code == code)
            ?? new OpeningStatusDefinition(code, "Неизвестный статус", "#BBC3D1");
    }

    public sealed class OpeningStatusDefinition
    {
        public string Code { get; }
        public string Name { get; }
        public string Color { get; }

        internal OpeningStatusDefinition(string code, string name, string color)
        {
            Code = code;
            Name = name;
            Color = color;
        }
    }


    internal sealed class OpeningSnapshot
    {
        // Имена полей одинаковы для G1 и G2. Линейные величины храним в миллиметрах.
        internal static readonly string[] Fields =
        {
            "Status", "StatusAuthor", "StatusLogin", "ConfirmedDate", "ConfirmedAuthor", "ConfirmedLogin",
            "LevelId", "LevelName", "LevelElevation", "X", "Y", "Z", "Rotation", "FacingFlipped", "HandFlipped",
            "TypeId", "TypeName", "HostId", "DimensionsXml",
        };

        private readonly Dictionary<string, string> _values = new Dictionary<string, string>();

        internal string this[string field]
        {
            get => _values.TryGetValue(field, out string value) ? value : string.Empty;
            set => _values[field] = value ?? string.Empty;
        }

        internal OpeningSnapshot Copy()
        {
            var copy = new OpeningSnapshot();
            foreach (string field in Fields)
                copy[field] = this[field];

            return copy;
        }

        internal SortedDictionary<string, string> Dimensions
        {
            get
            {
                var result = new SortedDictionary<string, string>(StringComparer.Ordinal);
                if (string.IsNullOrEmpty(this["DimensionsXml"]))
                    return result;

                XElement root = XElement.Parse(this["DimensionsXml"]);
                if (root.Name != "Dimensions")
                    throw new FormatException("Неизвестный формат размеров отверстия");

                foreach (XElement item in root.Elements("Dimension"))
                    result.Add((string)item.Attribute("Name"), (string)item.Attribute("Value") ?? string.Empty);

                return result;
            }
            set => this["DimensionsXml"] = new XElement("Dimensions", value.Select(item => new XElement("Dimension",
                new XAttribute("Name", item.Key), new XAttribute("Value", item.Value)))).ToString(SaveOptions.DisableFormatting);
        }
    }


    public sealed class OpeningHistoryEntry
    {
        internal const string ChangesDetected = "Обнаружены изменения";
        internal const string ConfirmedByUser = "Подтверждено пользователем";
        internal const string ConfirmedByBim = "Подтверждено BIM-отделом";

        public string Date { get; set; } = string.Empty;
        public string Author { get; set; } = string.Empty;
        public string Action { get; set; } = string.Empty;
        public string Details { get; set; } = string.Empty;
        public bool HasDate => !string.IsNullOrEmpty(Date);
        public bool HasAuthor => !string.IsNullOrEmpty(Author);
    }

    internal sealed class OpeningHistoryRecord
    {
        internal const string UnknownAuthor = "Не определён";
        internal const int HistoryLimit = 5;
        internal string ElementUniqueId { get; set; } = string.Empty;
        internal string CreatedAuthor { get; set; } = string.Empty;
        internal string CreatorLogin { get; set; } = string.Empty;
        internal OpeningSnapshot Current { get; set; }
        internal OpeningSnapshot Previous { get; set; }
        internal long CompressedCount { get; private set; }
        internal List<OpeningHistoryEntry> Entries { get; } = new List<OpeningHistoryEntry>();

        internal void Append(OpeningHistoryEntry entry)
        {
            Entries.Add(entry);
            Compact();
        }

        private void Compact()
        {
            int excess = Entries.Count - HistoryLimit;
            if (excess <= 0)
                return;

            CompressedCount += excess;
            Entries.RemoveRange(0, excess);
        }

        internal OpeningHistoryEntry[] DisplayEntries()
        {
            IEnumerable<OpeningHistoryEntry> result = Entries.AsEnumerable().Reverse();
            if (CompressedCount > 0)
                result = result.Concat(new[] { new OpeningHistoryEntry
                {
                    Action = "Остальная история сжата",
                    Details = "Предыдущих записей: " + CompressedCount,
                } });

            return result.ToArray();
        }

        internal OpeningHistoryEntry[] DisplayChanges()
        {
            // В карточке показываем только различия поколений. Записи о создании
            // и выполнении команд сохраняются в хранилище, но не засоряют этот раздел.
            var changes = new List<OpeningHistoryEntry>();
            foreach (OpeningHistoryEntry entry in DisplayEntries())
            {
                if (entry.Action == OpeningHistoryEntry.ChangesDetected || entry.Action == "Остальная история сжата")
                {
                    changes.Add(entry);
                    continue;
                }

                if (entry.Action != OpeningHistoryEntry.ConfirmedByUser && entry.Action != OpeningHistoryEntry.ConfirmedByBim)
                    continue;

                // У ранее записанных подтверждений после перехода статуса мог идти
                // полный снимок. Он относится к геометрии, а не к различиям поколений.
                string transition = (entry.Details ?? string.Empty).Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .FirstOrDefault(line => line.StartsWith("Статус: ", StringComparison.Ordinal));
                if (string.IsNullOrEmpty(transition))
                    continue;

                changes.Add(new OpeningHistoryEntry
                {
                    Action = entry.Action,
                    Details = transition,
                });
            }

            return changes.ToArray();
        }

        internal void ReadHistory(string xml)
        {
            if (string.IsNullOrEmpty(xml))
                return;

            XElement root = XElement.Parse(xml);
            if (root.Name != "History")
                throw new FormatException("Неизвестный формат истории отверстия");

            CompressedCount = (long?)root.Attribute("CompressedCount") ?? 0;
            foreach (XElement item in root.Elements("Entry"))
                Entries.Add(new OpeningHistoryEntry
                {
                    Date = (string)item.Element("Date") ?? string.Empty,
                    Author = (string)item.Element("Author") ?? string.Empty,
                    Action = (string)item.Element("Action") ?? string.Empty,
                    Details = (string)item.Element("Details") ?? string.Empty,
                });

            Compact();
        }

        internal string HistoryXml() => new XElement("History", new XAttribute("CompressedCount", CompressedCount),
            Entries.Select(item => new XElement("Entry", new XElement("Date", item.Date ?? string.Empty),
                new XElement("Author", item.Author ?? string.Empty), new XElement("Action", item.Action ?? string.Empty),
                new XElement("Details", item.Details ?? string.Empty)))).ToString(SaveOptions.DisableFormatting);
    }


    public sealed class OpeningHistoryRow
    {
        internal Autodesk.Revit.DB.ElementId RevitId { get; set; }
        public string UniqueId { get; set; }
        public string ElementId { get; set; }
        public string TypeName { get; set; }
        public string LevelId { get; set; }
        public string LevelName { get; set; }
        public double LevelElevation { get; set; }
        public string GeometryDetails { get; set; }
        public string CreatedAuthor { get; set; }
        public string StatusCode { get; set; } = OpeningStatuses.New;
        public string Status => OpeningStatuses.Get(StatusCode).Name;
        public string StatusColor => OpeningStatuses.Get(StatusCode).Color;
        public string StatusAuthor { get; set; }
        public string ConfirmedDate { get; set; }
        public string ConfirmedDateDisplay => string.IsNullOrEmpty(ConfirmedDate) ? "Не подтверждено" : ConfirmedDate;
        public string ConfirmedAuthor { get; set; }
        public string CacheNote { get; set; }
        public OpeningHistoryEntry[] Entries { get; set; }
        public bool HasHistory => Entries != null && Entries.Length > 0;

        public bool IsDataExpanded { get; set; }
        public bool IsGeometryExpanded { get; set; }
        public bool IsHistoryExpanded { get; set; }
        public bool IsConfirmed => StatusCode == OpeningStatuses.Confirmed;
        public string SearchText => ElementId + " " + TypeName + " " + CreatedAuthor + " " + ConfirmedAuthor
            + " " + StatusAuthor + " " + Status + " " + LevelName;
    }
}
