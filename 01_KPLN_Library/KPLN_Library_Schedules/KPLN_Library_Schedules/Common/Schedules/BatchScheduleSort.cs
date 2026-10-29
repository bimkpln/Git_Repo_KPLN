using System.Collections.Generic;

namespace KPLN_Library_Schedules.Common.Schedules
{
    public sealed class BatchScheduleSort : ScheduleObservableObject
    {
        public bool IsExisting { get; set; }
        public bool CanChangeParameter => true;
        public string InitialParameterKey { get; set; }
        public bool ParameterChanged => IsExisting && Parameter?.Key != InitialParameterKey;
        public Dictionary<string, int> SourceIndices { get; } = new Dictionary<string, int>();
        public string InitialSignature { get; set; }
        public string InitialPosition { get; set; }
        public string InitialVisibility { get; set; }
        public string Signature => Parameter?.Key + "\n" + Position + "\n" + Direction + "\n" + Visibility;
        public bool HasChanges => Enabled && (!IsExisting || Deleted || Signature != InitialSignature);
        public bool VisibilityChanged => Enabled && !Deleted && Visibility != "Исходное" && (!IsExisting || ParameterChanged || Visibility != InitialVisibility);
        public bool CanEdit => !Deleted;
        public string State => Deleted ? "Будет удалён" : IsExisting ? "Существующий" : "Новый";
        private bool _enabled = true;
        private bool _deleted;
        private ScheduleParameterChoice _parameter;
        private string _position = "Исходное";
        private string _direction = "По возрастанию";
        private string _visibility = "Исходное";
        public bool Enabled { get => _enabled; set => Change(ref _enabled, value); }
        public bool Deleted
        {
            get => _deleted;
            set { Change(ref _deleted, value); Notify(nameof(State)); Notify(nameof(CanEdit)); }
        }
        public ScheduleParameterChoice Parameter { get => _parameter; set => Change(ref _parameter, value); }
        public string Position { get => _position; set => Change(ref _position, value); }
        public string Direction { get => _direction; set => Change(ref _direction, value); }
        public string Visibility { get => _visibility; set => Change(ref _visibility, value); }
        public static readonly string[] Directions = { "Исходное", "По возрастанию", "По убыванию" };
    }
}
