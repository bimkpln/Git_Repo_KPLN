using Newtonsoft.Json;
using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_Tools.Common.OVVK_System
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
                if (_parameterName != value)
                {
                    _parameterName = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public string PartOfInsulationName
        {
            get => _partOfInsulationName;
            set
            {
                if (_partOfInsulationName != value)
                {
                    _partOfInsulationName = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public string PartOfSystemName
        {
            get => _partOfSystemName;
            set
            {
                if (_partOfSystemName != value)
                {
                    _partOfSystemName = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public object ToJson()
        {
            return new
            {
                this.ParameterName,
                this.PartOfInsulationName,
                this.PartOfSystemName
            };
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
