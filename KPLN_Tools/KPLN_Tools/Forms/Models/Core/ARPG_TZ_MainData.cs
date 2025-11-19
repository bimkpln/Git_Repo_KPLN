using Autodesk.Revit.DB;
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

        private string _flatType = "Room";
        private string _flatAreaCoeff = "0,95";
        private string _logAreaCoeff = "0,5";
        private string _balkAreaCoeff = "0,3";
        private string _terraceAreaCoeff = "0,3";
        private bool _gripCorpParam = false;
        private bool _gripSectParam = false;

        public ARPG_TZ_MainData() { }

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
        /// Группировать помещения по корпусу
        /// </summary>
        public bool IsGripCorpParam
        {
            get => _gripCorpParam;
            set
            {
                _gripCorpParam = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Группировать помещения по секции
        /// </summary>
        public bool IsGripSectParam
        {
            get => _gripSectParam;
            set
            {
                _gripSectParam = value;
                NotifyPropertyChanged();
            }
        }

        public BuiltInParameter FlatAreaParam { get; set; }

        /// <summary>
        /// Имя параметра для Имя помещения (квартира\балкон\лоджия\терраса)
        /// </summary>
        public string FlatNameParamName { get; } = "Назначение";

        /// <summary>
        /// Имя параметра для Номера квартиры
        /// </summary>
        public string FlatNumbParamName { get; } = "ПОМ_Номер квартиры";

        /// <summary>
        /// Имя параметра для Номера этажа
        /// </summary>
        public string FlatLvlNumbParamName { get; } = "ПОМ_Номер этажа";

        /// <summary>
        /// Имя параметра для захваток (старт)
        /// </summary>
        public string GripCorpParamName { get; } = "ПОМ_Корпус";

        /// <summary>
        /// Имя параметра для захваток (после 1)
        /// </summary>
        public string GripSectParamName { get; } = "ПОМ_Секция";

        /// <summary>
        /// Погрешность при назначении кода квартиры в м2 (ВОЗМОЖНО БУДЕТ ПЕРЕВОД НА КОЭФФИЦИЕНТ)
        /// </summary>
        public string FlatAreaTolerance { get; set; } = "1";

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
        public string TZCodeParamName { get; } = "КВ_Код";

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
