using Autodesk.Revit.DB;
using System.Collections.Generic;
using System;
using System.Linq;

namespace KPLN_ViewsAndLists_Ribbon.Common.Schedules
{
    public sealed class ScheduleParameterChoice
    {
        public string Key { get; set; }
        public string Name { get; set; }
        public int ScheduleCount { get; set; }
        public string SuggestedKind { get; set; }
        public string Label => Name + " [" + ScheduleCount + " спец.; " + Key + "]";
    }

    public sealed class BatchScheduleFilter : ScheduleObservableObject
    {
        public bool IsAutomatic { get; set; }
        public bool CanToggle => !IsAutomatic;
        public bool IsExisting { get; set; }
        public bool CanChangeParameter => !IsExisting && !IsAutomatic;
        public Dictionary<string, int> SourceIndices { get; } = new Dictionary<string, int>();
        public string InitialSignature { get; set; }
        public string Signature => Condition + "\n" + Kind + "\n" + Value;
        public bool HasChanges => Enabled && (!IsExisting || Deleted || Signature != InitialSignature);
        private bool _deleted;
        public bool Deleted
        {
            get => _deleted;
            set
            {
                Change(ref _deleted, value);
                Notify(nameof(State));
                Notify(nameof(CanEdit));
            }
        }
        public bool CanEdit => !Deleted && !IsAutomatic;
        public string State => IsAutomatic ? "По значению копии" : Deleted ? "Будет удалён" : IsExisting ? "Существующий" : "Новый";
        private bool _enabled = true;
        private ScheduleParameterChoice _parameter;
        private string _condition = "Равно";
        private string _kind = "Не определён";
        private string _value = "";
        public bool Enabled { get => _enabled; set => Change(ref _enabled, value); }
        public ScheduleParameterChoice Parameter
        {
            get => _parameter;
            set
            {
                Change(ref _parameter, value);
                Kind = value?.SuggestedKind ?? "Не определён";
            }
        }
        public string Condition { get => _condition; set => Change(ref _condition, value); }
        public string Kind { get => _kind; set => Change(ref _kind, value); }
        public string Value { get => _value; set => Change(ref _value, value); }
        public static readonly string[] Kinds = { "Текст", "Целое число", "Число в единицах проекта", "ID элемента" };
        public static readonly string[] Conditions = {
            "Равно", "Не равно", "Больше", "Больше или равно", "Меньше", "Меньше или равно",
            "Содержит", "Не содержит", "Начинается с", "Не начинается с", "Заканчивается на", "Не заканчивается на",
            "Параметр существует", "Есть значение", "Нет значения", "Связан с глобальным параметром", "Не связан с глобальным параметром"
        };
    }

    public sealed class ScheduleOperation : ScheduleObservableObject
    {
        public bool? Committed { get; set; }
        public int WarningCount { get; set; }
        private string _result;
        public ScheduleItem Source { get; set; }
        public ScheduleLevelRow Level { get; set; }
        public ScheduleParameterChoice CopyParameter { get; set; }
        public ScheduleCopyValue CopyValue { get; set; }
        public int Index { get; set; }
        public bool CreateCopy { get; set; }
        public string TargetName { get; set; }
        public string SourceName => Source.Name;
        public string LevelName => Level.Name;
        public string ParameterName => CopyParameter?.Name ?? "—";
        public string ValueName => CopyValue?.Display ?? "—";
        public string Action => CreateCopy ? "Создать копию" : "Изменить выбранную";
        public string Result { get => _result; set => Change(ref _result, value); }
    }

    public sealed class ScheduleExecutionReport
    {
        public int Total { get; private set; }
        public int Created { get; private set; }
        public int Modified { get; private set; }
        public int Failed { get; private set; }
        public int NotExecuted { get; private set; }
        public int ParameterCopies { get; private set; }
        public int Warnings { get; private set; }
        public string Duration { get; private set; }
        public string Error { get; private set; }
        public string Details { get; private set; }
        public string Summary => "Создано: " + Created + "; изменено: " + Modified
            + "; ошибок: " + Failed + "; не выполнено: " + NotExecuted + ". Время: " + Duration + ".";
        public string Text => "Запланировано операций: " + Total
            + "\n\nСоздано спецификаций: " + Created
            + "\nИз них по значениям параметра: " + ParameterCopies
            + "\nИзменено существующих спецификаций: " + Modified
            + "\nОпераций с ошибкой / откатом: " + Failed
            + "\nНе выполнено: " + NotExecuted
            + "\nПредупреждений Revit: " + Warnings
            + "\n\nВремя выполнения: " + Duration
            + (string.IsNullOrEmpty(Error) ? "" : "\n\n" + Error)
            + "\n\nВ числе созданных и изменённых учтены только сохранённые операции."
            + "\nОтчёт относится к последнему запуску; последующая отмена в Revit его не пересчитывает."
            + "\n\nПодробности операций:\n" + Details;

        public static ScheduleExecutionReport Create(IEnumerable<ScheduleOperation> operations, TimeSpan elapsed, string error = null)
        {
            var rows = operations.ToList();
            return new ScheduleExecutionReport
            {
                Total = rows.Count,
                Created = rows.Count(row => row.Committed == true && row.CreateCopy),
                Modified = rows.Count(row => row.Committed == true && !row.CreateCopy),
                Failed = rows.Count(row => row.Committed == false),
                NotExecuted = rows.Count(row => !row.Committed.HasValue),
                ParameterCopies = rows.Count(row => row.Committed == true && row.CreateCopy && row.CopyValue != null),
                Warnings = rows.Sum(row => row.WarningCount),
                Duration = elapsed.TotalSeconds < 60 ? elapsed.TotalSeconds.ToString("0.0") + " с"
                    : ((int)elapsed.TotalMinutes) + " мин " + elapsed.Seconds + " с",
                Error = error,
                Details = string.Join("\n\n", rows.Select((row, index) => (index + 1) + ". " + row.TargetName + "\n" + row.Result))
            };
        }
    }

    // Значения переживают откат пробной спецификации: никаких объектов Parameter/Element.
    public sealed class ScheduleCopyValue : ScheduleObservableObject
    {
        private bool _selected = false;
        public bool Selected { get => _selected; set => Change(ref _selected, value); }
        public StorageType Storage { get; set; }
        public string Text { get; set; }
        public bool IsPrefix { get; set; }
        public int Integer { get; set; }
        public double Number { get; set; }
        public long ElementIdValue { get; set; }
        public string Display { get; set; }
        public int ElementCount { get; set; }
        public string Key => Storage + ":" + (Storage == StorageType.String ? Text
            : Storage == StorageType.Integer ? Integer.ToString(System.Globalization.CultureInfo.InvariantCulture)
            : Storage == StorageType.Double ? Number.ToString("R", System.Globalization.CultureInfo.InvariantCulture)
            : ElementIdValue.ToString(System.Globalization.CultureInfo.InvariantCulture));
    }
}
