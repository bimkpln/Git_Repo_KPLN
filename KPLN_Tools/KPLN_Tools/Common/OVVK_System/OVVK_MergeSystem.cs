using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_Tools.Common.OVVK_System
{
    public class OVVK_MergeSystem : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _firstParamName;
        private string _lastParamName;

        public OVVK_MergeSystem(string[] systemParameters)
        {
            SystemParameters = systemParameters;
        }

        public string[] SystemParameters { get; private set; }

        /// <summary>
        /// Первая часть значения системы для склеивания
        /// </summary>
        public string FirstParamName 
        {
            get => _firstParamName;
            set
            {
                _firstParamName = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Вторая часть значения системы для склеивания
        /// </summary>
        public string LastParamName
        {
            get => _lastParamName;
            set
            {
                _lastParamName = value;
                OnPropertyChanged();
            }
        }

        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
