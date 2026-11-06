using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_OpeningHoleManager.Core
{
    public enum OpeningHistoryStatusFilter { All, Unconfirmed, Pending, Changed, Confirmed }

    public sealed class OpeningFilterOption
    {
        public string Id { get; set; }
        public string Name { get; set; }
    }

    public sealed class OpeningStatusOption
    {
        public OpeningHistoryStatusFilter Value { get; set; }
        public string Name { get; set; }
    }

    public sealed class OpeningHistoryFilterSettings : INotifyPropertyChanged
    {
        public const string AllLevels = "*";
        private string _levelId = AllLevels;
        private bool _onlyActiveView;
        private OpeningHistoryStatusFilter _status;

        public event PropertyChangedEventHandler PropertyChanged;
        // * — все уровни, пустая строка — отверстия без привязки к уровню.
        public string LevelId { get => _levelId; set { _levelId = value; Changed(); } }
        public bool OnlyActiveView { get => _onlyActiveView; set { _onlyActiveView = value; Changed(); } }
        public OpeningHistoryStatusFilter Status { get => _status; set { _status = value; Changed(); } }
        public OpeningFilterOption[] Levels { get; set; } = new OpeningFilterOption[0];
        public OpeningStatusOption[] StatusOptions { get; } = new[]
        {
            new OpeningStatusOption { Value = OpeningHistoryStatusFilter.All, Name = "Все статусы" },
            new OpeningStatusOption { Value = OpeningHistoryStatusFilter.Unconfirmed, Name = "Все неподтверждённые" },
            new OpeningStatusOption { Value = OpeningHistoryStatusFilter.Pending, Name = "Ожидает подтверждения" },
            new OpeningStatusOption { Value = OpeningHistoryStatusFilter.Changed, Name = "Изменено после подтверждения" },
            new OpeningStatusOption { Value = OpeningHistoryStatusFilter.Confirmed, Name = "Подтверждено" }
        };

        public int ActiveCount => (LevelId != AllLevels ? 1 : 0) + (OnlyActiveView ? 1 : 0)
            + (Status != OpeningHistoryStatusFilter.All ? 1 : 0);

        public OpeningHistoryFilterSettings Copy() => new OpeningHistoryFilterSettings
        {
            LevelId = LevelId, OnlyActiveView = OnlyActiveView, Status = Status, Levels = Levels
        };

        public void Reset()
        {
            LevelId = AllLevels;
            OnlyActiveView = false;
            Status = OpeningHistoryStatusFilter.All;
        }

        internal bool Matches(OpeningHistoryRow row, string search, ISet<string> visibleIds)
        {
            if (LevelId != AllLevels && !string.Equals(LevelId, row.LevelId ?? string.Empty, StringComparison.Ordinal))
                return false;
            if (OnlyActiveView && (visibleIds == null || !visibleIds.Contains(row.ElementId)))
                return false;

            switch (Status)
            {
                case OpeningHistoryStatusFilter.Unconfirmed:
                    if (row.IsConfirmed) return false;
                    break;
                case OpeningHistoryStatusFilter.Pending:
                    if (row.Status != "Ожидает подтверждения") return false;
                    break;
                case OpeningHistoryStatusFilter.Changed:
                    if (row.Status != "Изменено после подтверждения") return false;
                    break;
                case OpeningHistoryStatusFilter.Confirmed:
                    if (!row.IsConfirmed) return false;
                    break;
            }

            return string.IsNullOrWhiteSpace(search)
                || row.SearchText.IndexOf(search.Trim(), StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void Changed([CallerMemberName] string property = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property));
    }
}
