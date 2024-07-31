using Autodesk.Revit.DB.Plumbing;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_Tools.Common.OVVK_System
{
    [Serializable]
    public class PipeThicknessEntity : INotifyPropertyChanged, IJsonSerializable
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _currentPipeTypeName;

        [JsonConstructor]
        public PipeThicknessEntity()
        {
        }

        public PipeThicknessEntity(PipeType pipeType)
        {
            CurrentPipeType = pipeType;
            CurrentPipeTypeName = pipeType.Name;
        }

        public PipeType CurrentPipeType { get; set; }

        public List<PipeTypeDiamAndThickness> CurrentPipeTypeDiamAndThickness { get; set; }

        public string CurrentPipeTypeName
        {
            get => _currentPipeTypeName;
            set
            {
                if (_currentPipeTypeName != value)
                {
                    _currentPipeTypeName = value;
                    NotifyPropertyChanged();
                }
            }
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public object ToJson()
        {
            return new
            {
                this.CurrentPipeTypeName,
                this.CurrentPipeTypeDiamAndThickness
            };
        }
    }
}
