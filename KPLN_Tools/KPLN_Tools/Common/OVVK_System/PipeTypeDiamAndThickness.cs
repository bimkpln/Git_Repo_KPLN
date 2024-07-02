using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_Tools.Common.OVVK_System
{
    [Serializable]
    public class PipeTypeDiamAndThickness : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private double _currentDiameter = 0.0;
        private double _currentThickness = 0.0;

        public PipeTypeDiamAndThickness(double currentDiameter)
        {
            CurrentDiameter = currentDiameter;
        }

        public double CurrentDiameter
        {
            get => _currentDiameter;
            private set
            {
                if (_currentDiameter != value)
                {
                    _currentDiameter = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public double CurrentThickness
        {
            get => _currentThickness;
            set
            {
                _currentThickness = value;
                NotifyPropertyChanged();
            }
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
