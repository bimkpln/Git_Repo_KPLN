using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_Loader.Forms.Common
{
    internal enum MainStatus
    {
        Envirnment,
        DbConnection,
        ModulesActivation
    }

    /// <summary>
    /// Описание статусов для основных маркеров событий для окна LoaderStatusForm
    /// </summary>
    internal class LoaderStatusEntity: INotifyPropertyChanged
    {
        private string _strStatus;
        private string _currentToolTip  = "Не реализовано!";
        private System.Windows.Media.Brush _currentStrStatusColor = System.Windows.Media.Brushes.Red; 

        public LoaderStatusEntity(string descr, string sts, MainStatus mainStatus)
        {
            Description = descr;
            StrStatus = sts;
            CurrentMainStatus = mainStatus;
        }
        
        public string Description { get; }
        
        public string StrStatus 
        { 
            get => _strStatus;
            set
            {
                _strStatus = value;
                NotifyPropertyChanged();
            }
        }
        
        public System.Windows.Media.Brush CurrentStrStatusColor
        {
            get => _currentStrStatusColor;
            set
            {
                _currentStrStatusColor = value;
                NotifyPropertyChanged();
            }
        }

        public MainStatus CurrentMainStatus { get; }

        public string CurrentToolTip
        {
            get => _currentToolTip;
            set
            {
                _currentToolTip = value;
                NotifyPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
