using ElementId = Autodesk.Revit.DB.ElementId;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace KPLN_OpeningHoleManager.Core
{
    internal static class OpeningHistoryDateDisplay
    {
        internal static string Format(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return OpeningHistoryRecord.UnknownDate;

            // Скрываем смещение в интерфейсе, сохраняя исходное местное время и данные хранилища.
            DateTimeOffset date;
            return DateTimeOffset.TryParseExact(value, "dd.MM.yyyy HH:mm:ss zzz", CultureInfo.InvariantCulture,
                DateTimeStyles.None, out date)
                ? date.ToString("dd.MM.yyyy HH:mm:ss", CultureInfo.InvariantCulture)
                : value;
        }
    }

    internal sealed class OpeningCreationInfo
    {
        public string Date { get; set; }
        public string CreatorLogin { get; set; }
    }

    public sealed class OpeningHistoryEntry
    {
        public string Date { get; set; }
        public string Author { get; set; }
        public string Action { get; set; }
        public string Details { get; set; }
        public string DisplayDate => OpeningHistoryDateDisplay.Format(Date);
    }

    internal sealed class OpeningHistoryRecord
    {
        internal const string UnknownDate = "Не определена";
        internal const string UnknownAuthor = "Не определён";
        public string CreatedDate { get; set; }
        public string CreatedAuthor { get; set; }
        public string CreatorLogin { get; set; }
        public string ElementUniqueId { get; set; }
        public string ConfirmedFingerprint { get; set; }
        public List<OpeningHistoryEntry> Entries { get; set; } = new List<OpeningHistoryEntry>();

        internal static List<OpeningHistoryEntry> Parse(string xml)
        {
            if (string.IsNullOrEmpty(xml))
                return new List<OpeningHistoryEntry>();
            XElement root = XElement.Parse(xml);
            if (root.Name != "History")
                throw new System.FormatException("Неизвестный формат истории отверстия");
            return root.Elements("Entry").Select(e => new OpeningHistoryEntry
            {
                Date = (string)e.Element("Date") ?? UnknownDate,
                Author = (string)e.Element("Author") ?? UnknownAuthor,
                Action = (string)e.Element("Action") ?? string.Empty,
                Details = (string)e.Element("Details") ?? string.Empty
            }).ToList();
        }

        internal string ToXml() => new XElement("History", Entries.Select(e => new XElement("Entry",
            new XElement("Date", e.Date ?? UnknownDate),
            new XElement("Author", e.Author ?? UnknownAuthor),
            new XElement("Action", e.Action ?? string.Empty),
            new XElement("Details", e.Details ?? string.Empty)))).ToString(SaveOptions.DisableFormatting);
    }

    public sealed class OpeningHistoryRow
    {
        // Только идентификаторы: карточка не удерживает устаревающие native-элементы Revit.
        internal ElementId RevitId { get; set; }
        internal string[] DependencyIds { get; set; } = new string[0];


        public string UniqueId { get; set; }
        public string ElementId { get; set; }
        public string TypeName { get; set; }
        public string LevelId { get; set; }
        public string LevelName { get; set; }
        public double LevelElevation { get; set; }
        public string CreatedDate { get; set; }
        public string CreatedAuthor { get; set; }
        public string Status { get; set; }
        public string CacheNote { get; set; }
        public OpeningHistoryEntry[] Entries { get; set; }
        public bool IsHistoryExpanded { get; set; }
        public bool IsConfirmed => Status == "Подтверждено";


        // Entries поступают из сервиса от новых событий к старым.
        private OpeningHistoryEntry LastConfirmation => Entries?.FirstOrDefault(entry =>
            entry.Action == "Подтверждено пользователем" || entry.Action == "Подтверждено BIM-отделом");
        public string CreatedDateDisplay => OpeningHistoryDateDisplay.Format(CreatedDate);
        public string ConfirmedDateDisplay => LastConfirmation != null
            ? OpeningHistoryDateDisplay.Format(LastConfirmation.Date)
            : IsConfirmed ? OpeningHistoryRecord.UnknownDate : "Не подтверждено";
        public string ConfirmedAuthor => LastConfirmation == null ? "—"
            : string.IsNullOrWhiteSpace(LastConfirmation.Author) ? OpeningHistoryRecord.UnknownAuthor : LastConfirmation.Author;
        public string SearchText => ElementId + " " + TypeName + " " + CreatedAuthor + " " + ConfirmedAuthor + " " + Status + " " + LevelName;
    }
}
