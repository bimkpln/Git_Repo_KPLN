using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_Tools.Common
{
    public class MonitorParamRule : INotifyPropertyChanged
    {
        public ObservableCollection<string> DocParamColl { get; set; }
        public ObservableCollection<string> LinkParamColl { get; set; }
        public event PropertyChangedEventHandler PropertyChanged;

        private string _selectedSourceParameter;
        private string _selectedTargetParameter;

        public MonitorParamRule(ObservableCollection<string> docParamColl, ObservableCollection<string> linkParamColl)
        {
            DocParamColl = docParamColl;
            LinkParamColl = linkParamColl;
        }

        public string SelectedSourceParameter 
        {
            get => _selectedSourceParameter;
            set
            {
                _selectedSourceParameter = value;
                NotifyPropertyChanged();
            }
        }

        public string SelectedTargetParameter
        {
            get => _selectedTargetParameter;
            set
            {
                _selectedTargetParameter = value;
                NotifyPropertyChanged();
            }
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
