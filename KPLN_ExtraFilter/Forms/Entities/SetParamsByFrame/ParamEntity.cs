using Autodesk.Revit.DB;
using Newtonsoft.Json;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_ExtraFilter.Forms.Entities
{
    /// <summary>
    /// Класс-сущность для параметра
    /// </summary>
    public class ParamEntity : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _currentName;
        private int _currentParamIntId;
        private string _currentToolTip;

        [JsonConstructor]
        public ParamEntity()
        {
        }

        public ParamEntity(Parameter param, string tooltip)
        {
            CurrentToolTip = tooltip;

            CurrentParamName = param.Definition.Name;
            CurrentParamIntId = param.Id.IntegerValue;
        }

        /// <summary>
        /// Имя параметра
        /// </summary>
        public string CurrentParamName
        {
            get => _currentName;
            set
            {
                _currentName = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// ID параметра
        /// </summary>
        public int CurrentParamIntId
        {
            get => _currentParamIntId;
            set
            {
                _currentParamIntId = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Дополнительное описание пар-ра для wpf
        /// </summary>
        public string CurrentToolTip
        {
            get => _currentToolTip;
            set
            {
                _currentToolTip = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Переопределение метода Equals. ОБЯЗАТЕЛЬНО для десериализации, т.к. иначе на wpf не может найти эквивалетный инстанс
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (obj is ParamEntity other)
            {
                return CurrentParamName == other.CurrentParamName && CurrentParamIntId == other.CurrentParamIntId;
            }
            return false;
        }

        /// <summary>
        /// Переопределение метода GetHashCode
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            // Используем простое XOR-сочетание хэш-кодов свойств
            return CurrentParamName.GetHashCode() ^ CurrentParamIntId.GetHashCode();
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
