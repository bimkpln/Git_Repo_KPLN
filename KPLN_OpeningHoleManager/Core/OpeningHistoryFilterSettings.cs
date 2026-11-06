using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_OpeningHoleManager.Core
{
    public sealed class OpeningFilterOption
    {
        public string Id { get; set; }
        public string Name { get; set; }
    }

    public sealed class OpeningStatusOption
    {
        public string Value { get; set; }
        public string Name { get; set; }
    }

    public sealed class OpeningHistoryFilterSettings : INotifyPropertyChanged
    {
        public const string AllLevels = "*";
        public const string AllStatuses = "*";
        public const string Unconfirmed = "unconfirmed";
        private string _levelId = AllLevels;
        private bool _onlyActiveView;
        private string _status = AllStatuses;

        public event PropertyChangedEventHandler PropertyChanged;
        // * — все уровни, пустая строка — отверстия без привязки к уровню.
        public string LevelId { get => _levelId; set { _levelId = value; Changed(); } }
        public bool OnlyActiveView { get => _onlyActiveView; set { _onlyActiveView = value; Changed(); } }
        public string Status { get => _status; set { _status = value; Changed(); } }
        public OpeningFilterOption[] Levels { get; set; } = new OpeningFilterOption[0];
        public OpeningStatusOption[] StatusOptions { get; } = new[]
        {
            new OpeningStatusOption { Value = AllStatuses, Name = "Все статусы" },
            new OpeningStatusOption { Value = Unconfirmed, Name = "Все неподтверждённые" },
        }.Concat(OpeningStatuses.All.Select(status => new OpeningStatusOption { Value = status.Code, Name = status.Name })).ToArray();

        public int ActiveCount => (LevelId != AllLevels ? 1 : 0) + (OnlyActiveView ? 1 : 0)
            + (Status != AllStatuses ? 1 : 0);

        public OpeningHistoryFilterSettings Copy() => new OpeningHistoryFilterSettings
        {
            LevelId = LevelId, OnlyActiveView = OnlyActiveView, Status = Status, Levels = Levels
        };

        public void Reset()
        {
            LevelId = AllLevels;
            OnlyActiveView = false;
            Status = AllStatuses;
        }

        internal bool Matches(OpeningHistoryRow row, string search, ISet<string> visibleIds)
        {
            if (LevelId != AllLevels && !string.Equals(LevelId, row.LevelId ?? string.Empty, StringComparison.Ordinal))
                return false;
            if (OnlyActiveView && (visibleIds == null || !visibleIds.Contains(row.ElementId)))
                return false;

            if (Status == Unconfirmed && row.IsConfirmed)
                return false;
            if (Status != AllStatuses && Status != Unconfirmed && Status != row.StatusCode)
                return false;

            return string.IsNullOrWhiteSpace(search)
                || row.SearchText.IndexOf(search.Trim(), StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void Changed([CallerMemberName] string property = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property));
    }
}
