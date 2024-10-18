using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_Tools.Forms.Models
{
    public class OVVK_SystemManager_ViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _parameterName;
        private string _sysNameSeparator;

        /// <summary>
        /// Имя параметра в окне
        /// </summary>
        public string ParameterName
        {
            get => _parameterName;
            set
            {
                if (_parameterName != value)
                {
                    _parameterName = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Символ разделителя
        /// </summary>
        public string SysNameSeparator
        {
            get => _sysNameSeparator;
            set
            {
                if (_sysNameSeparator != value)
                {
                    _sysNameSeparator = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Коллекция групп систем по выбранному параметру (ДОРАБОТАТЬ, ПОКА - ЗАГЛУШКА)
        /// </summary>
        public ObservableCollection<string> SystemSumParameters { get; set; }

        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
