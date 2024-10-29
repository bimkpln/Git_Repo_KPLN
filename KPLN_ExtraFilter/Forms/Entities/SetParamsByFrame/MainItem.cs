using KPLN_ExtraFilter.Common;
using Newtonsoft.Json;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_ExtraFilter.Forms.Entities
{
    /// <summary>
    /// Класс-сущность для WPF окна. Им комплектуется ItemsControl
    /// </summary>
    public class MainItem : INotifyPropertyChanged, IJsonSerializable
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _userInputParamValue;

        [JsonConstructor]
        public MainItem()
        {
        }

        public MainItem(ParamEntity userAddedEntity)
        {
            UserSelectedParamEntity = userAddedEntity;
        }

        /// <summary>
        /// Выбранный пользователем параметр
        /// </summary>
        public ParamEntity UserSelectedParamEntity { get; set; }

        /// <summary>
        /// Введенное пользователем значение параметра
        /// </summary>
        public string UserInputParamValue
        {
            get => _userInputParamValue;
            set
            {
                _userInputParamValue = value;
                NotifyPropertyChanged();
            }
        }

        public object ToJson() => new
        {
            // Parameter - не стоит добавлять в JSON, переваривается плохо.
            // Нужно на чтении JSON уточнять значения по CurrentParam
            this.UserSelectedParamEntity,
            this.UserInputParamValue,
        };

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

}
