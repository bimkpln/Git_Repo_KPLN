using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_Tools.Forms.Models.Core
{
    /// <summary>
    /// Контейнер типов квартир по ТЗ 
    /// </summary>
    public class ARPG_TZ_FlatData : INotifyPropertyChanged, IDataErrorInfo
    {
        public event PropertyChangedEventHandler PropertyChanged;
        
        private string _flatRangeName = "<имя диап.>";
        private string _flatCode = "<код кв.>";
        private string _flatPercent = "<%>";
        private string _flatAreaMin = "<от...>";
        private string _flatAreaMax = "<...до>";

        public ARPG_TZ_FlatData() { }

        /// <summary>
        /// Имя диапазона
        /// </summary>
        public string TZRangeName
        {
            get => _flatRangeName;
            set
            {
                if (_flatRangeName != value)
                {
                    _flatRangeName = value;
                    NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Код квартиры
        /// </summary>
        public string TZCode
        {
            get => _flatCode;
            set
            {
                if (_flatCode != value)
                {
                    _flatCode = value;
                    NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Процент по ТЗ
        /// </summary>
        public string TZPercent
        {
            get => _flatPercent;
            set
            {
                if (_flatPercent != value)
                {
                    _flatPercent = value;
                    NotifyPropertyChanged();

                    if (double.TryParse(_flatPercent, out double perc))
                        TZPercent_Double = perc;
                }
            }
        }

        /// <summary>
        /// Процент по ТЗ, тип Double
        /// </summary>
        public double TZPercent_Double { get; set; }

        /// <summary>
        /// Диапазон по ТЗ - минимум
        /// </summary>
        public string TZAreaMin
        {
            get => _flatAreaMin;
            set
            {
                if (_flatAreaMin != value)
                {
                    _flatAreaMin = value;
                    NotifyPropertyChanged();

                    if (double.TryParse(_flatAreaMin, out double min))
                        TZAreaMin_Double = min;
                }
            }
        }

        /// <summary>
        /// Диапазон по ТЗ - минимум, тип Double
        /// </summary>
        public double TZAreaMin_Double { get; set; }

        /// <summary>
        /// Диапазон по ТЗ - максимум
        /// </summary>
        public string TZAreaMax
        {
            get => _flatAreaMax;
            set
            {
                if (_flatAreaMax != value)
                {
                    _flatAreaMax = value;
                    NotifyPropertyChanged();

                    if (double.TryParse(_flatAreaMax, out double max))
                        TZAreaMax_Double = max;
                }
            }
        }

        /// <summary>
        /// Диапазон по ТЗ - максимум, тип Double
        /// </summary>
        public double TZAreaMax_Double { get; set; }

        public string Error => null;

        public string this[string columnName]
        {
            get
            {
                switch (columnName)
                {
                    case nameof(TZPercent):
                        return ValidatePercent(TZPercent);
                    case nameof(TZAreaMax):
                        return ValidateDouble(TZAreaMax);
                }

                return null;
            }
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private string ValidateDouble(string value)
        {
            if (string.IsNullOrEmpty(value))
                return null;

            return !double.TryParse(value, out double res) ? $"Можно вводить ТОЛЬКО числа больше 0, в формате \"0,0\"" : null;
        }

        private string ValidatePercent(string value)
        {
            if (!string.IsNullOrEmpty(value)
                && double.TryParse(value, out double res)
                && (res > 0 && res < 100))
                return null;

            return $"Процент выбираем из диапазона 0...100, в формате \"0,0\"";
        }
    }
}
