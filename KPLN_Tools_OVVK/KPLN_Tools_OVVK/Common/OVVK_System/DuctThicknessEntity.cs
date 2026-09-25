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
        internal const string ConfigName = "OV_DuctThickness";

        public event PropertyChangedEventHandler PropertyChanged;

        private string _parameterName;
        private string _partOfInsulationName;
        private string _partOfSystemName;

        [JsonConstructor]
        public DuctThicknessEntity()
        {
            ParameterName = "КП_И_Толщина стенки";
            PartOfInsulationName = "EI";
            PartOfSystemTypeName = "ДУ~ПД";
        }

        /// <summary>
        /// Режим определяется по текущему документу, а не по перенесённому конфигу.
        /// </summary>
        [JsonIgnore]
        public bool UseProtectionParameters { get; internal set; }

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
                _partOfSystemName = value ?? string.Empty;
                NotifyPropertyChanged();
                
                PartsOfSystemTypeName = _partOfSystemName.Split('~');
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
