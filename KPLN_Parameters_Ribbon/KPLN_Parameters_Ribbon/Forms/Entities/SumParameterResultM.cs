using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_Parameters_Ribbon.Forms.Entities
{
    public sealed class SumParameterResultM : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public string ParameterName { get; set; }

        public string SearchParameterName { get; set; }

        public string UnitName { get; set; }

        public double Sum { get; set; }

        private string _valueText;
        public string ValueText
        {
            get => _valueText;
            set
            {
                _valueText = value;
                NotifyPropertyChanged();
            }
        }

        private string _valueWithCoefficientText;
        public string ValueWithCoefficientText
        {
            get => _valueWithCoefficientText;
            set
            {
                _valueWithCoefficientText = value;
                NotifyPropertyChanged();
            }
        }

        private string _valueWithoutCoefficientText;
        public string ValueWithoutCoefficientText
        {
            get => _valueWithoutCoefficientText;
            set
            {
                _valueWithoutCoefficientText = value;
                NotifyPropertyChanged();
            }
        }

        public void NotifyAll()
        {
            NotifyPropertyChanged(nameof(ParameterName));
            NotifyPropertyChanged(nameof(SearchParameterName));
            NotifyPropertyChanged(nameof(UnitName));
            NotifyPropertyChanged(nameof(ValueText));
            NotifyPropertyChanged(nameof(ValueWithCoefficientText));
            NotifyPropertyChanged(nameof(ValueWithoutCoefficientText));
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "") =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
