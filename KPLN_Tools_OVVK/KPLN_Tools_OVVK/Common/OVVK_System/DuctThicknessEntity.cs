using KPLN_Library_ConfigWorker.Core;
using Newtonsoft.Json;
using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_Tools_OVVK.Common.OVVK_System
{
    [Serializable]
    public class DuctThicknessEntity : INotifyPropertyChanged, IJsonSerializable
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _parameterName;
        private string _partOfInsulationName;
        private string _partOfSystemName;

        [JsonConstructor]
        public DuctThicknessEntity()
        {
        }

        public string ParameterName
        {
            get => _parameterName;
            set
            {
                _parameterName = value;
                NotifyPropertyChanged();
            }
        }

        public string PartOfInsulationName
        {
            get => _partOfInsulationName;
            set
            {
                _partOfInsulationName = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Общее поле для ввода
        /// </summary>
        public string PartOfSystemTypeName
        {
            get => _partOfSystemName;
            set
            {
                _partOfSystemName = value;
                NotifyPropertyChanged();
                
                string[] splitedSysName = _partOfSystemName.Split('~');
                if (splitedSysName.Length > 1)
                    PartsOfSystemTypeName = splitedSysName;
                else
                    PartsOfSystemTypeName = new string[1] { _partOfSystemName };
            }
        }

        /// <summary>
        /// Расчлененное на части имена систем
        /// </summary>
        public string[] PartsOfSystemTypeName { get; private set; }

        public object ToJson()
        {
            return new
            {
                this.ParameterName,
                this.PartOfInsulationName,
                this.PartOfSystemTypeName
            };
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
