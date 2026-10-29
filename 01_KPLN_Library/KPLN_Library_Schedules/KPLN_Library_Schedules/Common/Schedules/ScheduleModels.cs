using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace KPLN_Library_Schedules.Common.Schedules
{
    public class ScheduleLevelRow : ScheduleObservableObject
    {
        private bool _selected;
        private string _floor;
        private string _group = "";
        public Level Level { get; set; }
        public string Key => Level?.UniqueId ?? "single";
        public string Name => Level?.Name ?? "Без уровня";
        public double Height => (Level?.ProjectElevation ?? 0) * 0.3048;
        public string Elevation => ScheduleNaming.Elevation((Level?.Elevation ?? 0) * 0.3048);
        public string Existing { get; set; }
        public string Hint { get; set; }
        public bool Selected { get => _selected; set => Change(ref _selected, value); }
        public string Floor { get => _floor; set => Change(ref _floor, value); }
        public string Group { get => _group; set => Change(ref _group, value); }
    }

    public static class FilterModes
    {
        public const string Keep = "Сохранить";
        public const string Pattern = "По маске";
        public const string Level = "Целевой уровень";
        public const string Manual = "Вручную по этажам";
        public static readonly string[] All = { Keep, Pattern, Level, Manual };
    }

    public sealed class FilterRule : ScheduleObservableObject
    {
        private string _mode = FilterModes.Keep;
        private string _pattern = "{Этаж}";
        public int Index { get; set; }
        public ScheduleFilter Original { get; set; }
        public string FieldName { get; set; }
        public string FieldIdentity { get; set; }
        public string Condition
        {
            get
            {
                // Строковые имена сохраняют совместимость с разными версиями Revit API.
                switch (Original.FilterType.ToString())
                {
                    case "Equal": return "Равно";
                    case "NotEqual": return "Не равно";
                    case "GreaterThan": return "Больше";
                    case "GreaterThanOrEqual": return "Больше или равно";
                    case "LessThan": return "Меньше";
                    case "LessThanOrEqual": return "Меньше или равно";
                    case "Contains": return "Содержит";
                    case "NotContains": return "Не содержит";
                    case "BeginsWith": return "Начинается с";
                    case "NotBeginsWith": return "Не начинается с";
                    case "EndsWith": return "Заканчивается на";
                    case "NotEndsWith": return "Не заканчивается на";
                    case "HasParameter": return "Параметр существует";
                    case "HasValue": return "Есть значение";
                    case "HasNoValue": return "Нет значения";
                    case "IsAssociatedWithGlobalParameter": return "Связан с глобальным параметром";
                    case "IsNotAssociatedWithGlobalParameter": return "Не связан с глобальным параметром";
                    default: return "Неизвестное условие";
                }
            }
        }
        public string Value { get; set; }
        public string ValueType { get; set; }
        public string Mode { get => _mode; set => Change(ref _mode, value); }
        public string Pattern { get => _pattern; set => Change(ref _pattern, value); }
        public Dictionary<string, string> Overrides { get; } = new Dictionary<string, string>();
    }

    public class ScheduleItem : ScheduleObservableObject
    {
        private bool _copy = true;
        private bool _include = true;
        public ViewSchedule Schedule { get; set; }
        public ObservableCollection<FilterRule> Filters { get; } = new ObservableCollection<FilterRule>();
        // null — настройки старого формата; пустой список — явно настроенная спецификация без правил.
        public List<LevelScheduleFilter> LevelFilters { get; set; }
        public List<BatchScheduleSort> LevelSorts { get; set; }
        public string Name => Schedule.Name;
        public string Hint { get; set; }
        public bool Copy { get => _copy; set => Change(ref _copy, value); }
        public bool Include { get => _include; set => Change(ref _include, value); }
    }

    public sealed class FilterValueRow : ScheduleObservableObject
    {
        private string _value;
        public FilterRule Rule { get; set; }
        public ScheduleLevelRow Level { get; set; }
        public string Mode { get; set; }
        public bool CanEdit => Mode == FilterModes.Pattern || Mode == FilterModes.Manual;
        public string Explanation => Mode == FilterModes.Keep ? "Исходное значение; только просмотр"
            : Mode == FilterModes.Level ? "Выбранный уровень Revit; только просмотр"
            : Mode == FilterModes.Manual ? "Введите значение для этого этажа; пересчёт сохранит ввод"
            : "Результат маски; пересчёт сбросит ручные поправки";
        public string Field => Rule.FieldName + " / " + Rule.Condition;
        public string LevelName => Level.Name;
        public string Value { get => _value; set => Change(ref _value, value); }
    }

}
