using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_Tools.Forms.Models.Core
{
    /// <summary>
    /// Контейнер данных из ТЗ (общие для ВСЕХ квартир)
    /// </summary>
    public class ARPG_TZ_MainData : INotifyPropertyChanged, IDataErrorInfo
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _flatAreaCoeff = "0,95";
        private bool _heatingRoomsInPrj = true;
        private string _logAreaCoeff = "0,5";
        private string _balkAreaCoeff = "0,3";
        private string _terraceAreaCoeff = "0,3";
        private string _flatNameParamName = "Имя";
        private string _flatNumbParamName = "ПОМ_Номер квартиры";
        private string _flatLvlNumbParamName = "ПОМ_Номер этажа";
        private string _gripParamName1 = "ПОМ_Корпус";
        private string _gripParamName2;
        private string _flatType = "Room";

        public ARPG_TZ_MainData() { }

        /// <summary>
        /// Погрешность при назначении кода квартиры в м2 (ВОЗМОЖНО БУДЕТ ПЕРЕВОД НА КОЭФФИЦИЕНТ)
        /// </summary>
        public string FlatAreaTolerance { get; set; } = "1";

        /// <summary>
        /// Коэф. уменьшения квартир
        /// </summary>
        public string FlatAreaCoeff
        {
            get => _flatAreaCoeff;
            set
            {
                if (_flatAreaCoeff != value)
                {
                    _flatAreaCoeff = value;
                    NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// В проекте есть неотапливаемые помещения? 
        /// (влияет на необходимость привязки сепарированных помещений к основному)
        /// </summary>
        public bool HeatingRoomsInPrj
        {
            get => _heatingRoomsInPrj;
            set
            {
                if (_heatingRoomsInPrj != value)
                {
                    _heatingRoomsInPrj = value;
                    NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Коэф. уменьшения лоджий
        /// </summary>
        public string BalkAreaCoeff
        {
            get => _balkAreaCoeff;
            set
            {
                if (_balkAreaCoeff != value)
                {
                    _balkAreaCoeff = value;
                    NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Коэф. уменьшения терасс
        /// </summary>
        public string TerraceAreaCoeff
        {
            get => _terraceAreaCoeff;
            set
            {
                if (_terraceAreaCoeff != value)
                {
                    _terraceAreaCoeff = value;
                    NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Коэф. уменьшения лоджий
        /// </summary>
        public string LogAreaCoeff
        {
            get => _logAreaCoeff;
            set
            {
                if (_logAreaCoeff != value)
                {
                    _logAreaCoeff = value;
                    NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Имя параметра для Имя квартиры
        /// </summary>
        public string FlatNameParamName
        {
            get => _flatNameParamName;
            set
            {
                if (_flatNameParamName != value)
                {
                    _flatNameParamName = value;
                    NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Имя параметра для Номера квартиры
        /// </summary>
        public string FlatNumbParamName
        {
            get => _flatNumbParamName;
            set
            {
                if (_flatNumbParamName != value)
                {
                    _flatNumbParamName = value;
                    NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Имя параметра для Номера этажа
        /// </summary>
        public string FlatLvlNumbParamName
        {
            get => _flatLvlNumbParamName;
            set
            {
                if (_flatLvlNumbParamName != value)
                {
                    _flatLvlNumbParamName = value;
                    NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Имя параметра для захваток (старт)
        /// </summary>
        public string GripParamName1
        {
            get => _gripParamName1;
            set
            {
                if (_gripParamName1 != value)
                {
                    _gripParamName1 = value;
                    NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Имя параметра для захваток (после 1)
        /// </summary>
        public string GripParamName2
        {
            get => _gripParamName2;
            set
            {
                if (_gripParamName2 != value)
                {
                    _gripParamName2 = value;
                    NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Выбранный тип для рассчёта
        /// </summary>
        public string FlatType
        {
            get => _flatType;
            set
            {
                if (_flatType != value)
                {
                    _flatType = value;
                    NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Имя параметра для площади с коэффициентом
        /// </summary>
        public string AreaCoeffParamName { get; } = "ПОМ_Площадь_К";

        /// <summary>
        /// Имя параметра для площади с коэффициентом
        /// </summary>
        public string SumAreaCoeffParamName { get; } = "КВ_Площадь_Общая_К";

        /// <summary>
        /// Имя параметра для кода квартиры по ТЗ
        /// </summary>
        public string TZCodeParamName { get; } = "КВ_Диапазон_Тип квартиры";

        /// <summary>
        /// Имя параметра для имя диапазона по ТЗ
        /// </summary>
        public string TZRangeNameParamName { get; } = "КВ_Диапазон_Наименование";

        /// <summary>
        /// Имя параметра для процент диапазона по ТЗ 
        /// </summary>
        public string TZPercentParamName { get; } = "КВ_Диапазон_Процент по ТЗ";

        /// <summary>
        /// Имя параметра для диапазон min по ТЗ
        /// </summary>
        public string TZAreaMinParamName { get; } = "КВ_Диапазон min";

        /// <summary>
        /// Имя параметра для диапазон max по ТЗ
        /// </summary>
        public string TZAreaMaxParamName { get; } = "КВ_Диапазон max";

        /// <summary>
        /// Имя параметра для полученного процента в модели
        /// </summary>
        public string ModelPercentParamName { get; } = "КВ_Диапазон_Процент";

        /// <summary>
        /// Имя параметра для полученного отклонения от процента из ТЗ в модели
        /// </summary>
        public string ModelPercentToleranceParamName { get; } = "КВ_Диапазон_Процент_Отклонение от ТЗ";

        public string Error => null;

        public string this[string columnName]
        {
            get
            {
                switch (columnName)
                {
                    case nameof(FlatAreaTolerance):
                        return ValidateTolerance(FlatAreaTolerance);
                    case nameof(FlatAreaCoeff):
                        return ValidateCoeff(FlatAreaCoeff);
                    case nameof(BalkAreaCoeff):
                        return ValidateCoeff(BalkAreaCoeff);
                    case nameof(LogAreaCoeff):
                        return ValidateCoeff(LogAreaCoeff);
                }

                return null;
            }
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private string ValidateTolerance(string value)
        {
            if (!string.IsNullOrEmpty(value)
                && double.TryParse(value, out double res)
                && (res >= 0 && res <= 100))
                return null;

            return $"Коэффициент выбираем из диапазона 0...100, в формате \"0,0\"";
        }

        private string ValidateCoeff(string value)
        {
            if (!string.IsNullOrEmpty(value)
                && double.TryParse(value, out double res)
                && (res >= 0 && res <= 1))
                return null;

            return $"Коэффициент выбираем из диапазона 0...1, в формате \"0,0\"";
        }
    }
}
