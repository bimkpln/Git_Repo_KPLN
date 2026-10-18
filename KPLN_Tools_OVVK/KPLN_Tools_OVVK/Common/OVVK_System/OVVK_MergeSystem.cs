using System.ComponentModel;
using System.Runtime.CompilerServices;
using System;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Collections.ObjectModel;
using System.Collections.Generic;

namespace KPLN_Tools_OVVK.Common.OVVK_System
{
    public class OVVK_MergeSystem : INotifyPropertyChanged
    {
        public const int MinSystemCount = 2;
        public const int MaxSystemCount = 5;

        public static bool IsAllowedCount(int count) => count >= MinSystemCount && count <= MaxSystemCount;
        public event PropertyChangedEventHandler PropertyChanged;

        private bool _isRestoring;
        private static readonly Regex _systemCode = new Regex(@"^(\p{L}*)([0-9])([0-9]*(?:\.[0-9]+)*)$", RegexOptions.CultureInvariant);

        public OVVK_MergeSystem(string[] systemParameters)
        {
            SystemParameters = systemParameters;
            AddSystem();
            AddSystem();
        }

        public string[] SystemParameters { get; private set; }
        public ObservableCollection<OVVK_MergeSystemPart> Systems { get; } = new ObservableCollection<OVVK_MergeSystemPart>();
        public string[] SelectedSystemNames => Systems.Select(part => part.Value).ToArray();
        public bool CanAddSystem => Systems.Count < MaxSystemCount;
        public string AddSystemToolTip => $"Добавить систему в эту группу. Максимум: {MaxSystemCount}.";

        public void AddSystem()
        {
            if (!CanAddSystem)
                return;
            var part = new OVVK_MergeSystemPart(SystemParameters, Systems.Count == 0, Systems.Count >= MinSystemCount);
            part.PropertyChanged += OnPartChanged;
            Systems.Add(part);
            OnPropertyChanged(nameof(CanAddSystem));
        }

        public void RemoveSystem(OVVK_MergeSystemPart part)
        {
            if (Systems.IndexOf(part) < MinSystemCount)
                return;
            part.PropertyChanged -= OnPartChanged;
            Systems.Remove(part);
            OnPropertyChanged(nameof(CanAddSystem));
        }

        public void SetSystemNames(IEnumerable<string> names)
        {
            string[] values = names.ToArray();
            if (!IsAllowedCount(values.Length))
                throw new ArgumentException($"В группе должно быть от {MinSystemCount} до {MaxSystemCount} систем.", nameof(names));
            _isRestoring = true;
            try
            {
                while (Systems.Count > MinSystemCount)
                    RemoveSystem(Systems.Last());
                while (Systems.Count < values.Length)
                    AddSystem();
                for (int index = 0; index < Systems.Count; index++)
                    Systems[index].Value = index < values.Length ? values[index] : null;
            }
            finally { _isRestoring = false; }
        }

        private void OnPartChanged(object sender, PropertyChangedEventArgs args)
        {
            if (ReferenceEquals(sender, Systems[0]))
            {
                OnPropertyChanged(nameof(FirstParamName));
                // Автоподбор применяется только ко второй системе и не подменяет сохранённый конфиг.
                if (!_isRestoring)
                    LastParamName = FindSuggestedSystem(FirstParamName);
            }
            else if (ReferenceEquals(sender, Systems[1]))
                OnPropertyChanged(nameof(LastParamName));
            OnPropertyChanged(nameof(SelectedSystemNames));
        }

        /// <summary>
        /// Первая часть значения системы для склеивания
        /// </summary>
        public string FirstParamName 
        {
            get => Systems[0].Value;
            set => Systems[0].Value = value;
        }

        /// <summary>
        /// Вторая часть значения системы для склеивания
        /// </summary>
        public string LastParamName
        {
            get => Systems[1].Value;
            set => Systems[1].Value = value;
        }

        private string FindSuggestedSystem(string firstName)
        {
            if (string.IsNullOrEmpty(firstName))
                return null;

            Match match = _systemCode.Match(firstName);
            if (!match.Success)
                return null;

            int nextDigit = match.Groups[2].Value[0] - '0' + 1;
            string candidate = match.Groups[1].Value + nextDigit.ToString(CultureInfo.InvariantCulture) + match.Groups[3].Value;
            return SystemParameters?.FirstOrDefault(name => string.Equals(name, candidate, StringComparison.Ordinal));
        }

        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    public sealed class OVVK_MergeSystemPart : INotifyPropertyChanged
    {
        private string _value;
        public event PropertyChangedEventHandler PropertyChanged;
        public string[] SystemParameters { get; }
        public bool IsFirst { get; }
        public bool CanRemove { get; }

        internal OVVK_MergeSystemPart(string[] systems, bool isFirst, bool canRemove)
        {
            SystemParameters = systems;
            IsFirst = isFirst;
            CanRemove = canRemove;
        }

        public string Value
        {
            get => _value;
            set
            {
                if (_value == value)
                    return;
                _value = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Value)));
            }
        }
    }
}
