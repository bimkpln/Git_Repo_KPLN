using System.Collections.Generic;
using System.Linq;

namespace KPLN_Library_Schedules.Common.Schedules
{
    /// <summary>Общее правило конструктора с настройкой значений по этажам.</summary>
    public sealed class LevelScheduleFilter : BatchScheduleFilter
    {
        private string _lastParameterKey;
        private string _mode = FilterModes.Pattern;
        public string OriginalIdentity { get; set; }
        public string OriginalValue { get; set; }
        public string OriginalCondition { get; set; }
        internal string LevelParameterKey { get; set; }
        public bool IsLevelParameter
        {
            get => LevelParameterKey != null && LevelParameterKey == Parameter?.Key;
            set
            {
                LevelParameterKey = value ? Parameter?.Key : null;
                Notify(nameof(IsLevelParameter));
                Notify(nameof(Modes));
            }
        }
        public Dictionary<string, string> Overrides { get; } = new Dictionary<string, string>();
        public bool NeedsValue => Condition != "Параметр существует" && Condition != "Есть значение" && Condition != "Нет значения";
        public string Mode
        {
            get => _mode;
            set
            {
                Change(ref _mode, value);
                Notify(nameof(CanEditPattern));
                Notify(nameof(CanEditValue));
            }
        }
        public override bool CanEditValue => CanEditPattern;
        public bool CanEditPattern => CanEdit && Mode == FilterModes.Pattern && NeedsValue;
        public IEnumerable<string> Modes => FilterModes.All.Where(mode =>
            (IsExisting || mode != FilterModes.Keep) && (IsLevelParameter || mode != FilterModes.Level));
        public override bool HasChanges => FilterHasChanges || VisibilityChanged;
        public override bool FilterHasChanges => Enabled && (!IsExisting || Deleted || ParameterChanged
            || Condition != OriginalCondition || (NeedsValue && Mode != FilterModes.Keep
                && (Mode != FilterModes.Pattern || Value != OriginalValue || Overrides.Count > 0)));

        public LevelScheduleFilter()
        {
            PropertyChanged += (sender, args) =>
            {
                if (args.PropertyName == nameof(Value)) Overrides.Clear();
                if (args.PropertyName == nameof(Parameter) && _lastParameterKey != Parameter?.Key)
                {
                    _lastParameterKey = Parameter?.Key;
                    Overrides.Clear();
                    Notify(nameof(IsLevelParameter));
                    if (!IsLevelParameter && Mode == FilterModes.Level)
                        Mode = FilterModes.Pattern;
                    Notify(nameof(Modes));
                }
                if (args.PropertyName == nameof(Deleted) || args.PropertyName == nameof(Condition))
                {
                    Notify(nameof(CanEditPattern));
                    Notify(nameof(CanEditValue));
                }
            };
        }
    }

    public sealed class LevelFilterValueRow : ScheduleObservableObject
    {
        private string _value;
        public LevelScheduleFilter Rule { get; set; }
        public ScheduleLevelRow Level { get; set; }
        public string Mode { get; set; }
        public bool CanEdit { get; set; }
        public string Field { get; set; }
        public string LevelName => Level.Name;
        public string Value { get => _value; set => Change(ref _value, value); }
    }
}
