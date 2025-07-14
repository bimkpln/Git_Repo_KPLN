using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_BIMTools_Ribbon.Forms.Models
{
    /// <summary>
    /// Сущность для обработки и отображения в окне 
    /// </summary>
    public class DBRevitDocExchangesWrapper : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private bool _isSelected = false;

        public DBRevitDocExchangesWrapper(DBRevitDocExchanges dBRevitDocExchanges)
        {
            CurrentDBRevitDocExchanges = dBRevitDocExchanges;
        }

        public DBRevitDocExchanges CurrentDBRevitDocExchanges { get; }

        public int Id
        {
            get => CurrentDBRevitDocExchanges.Id;
            set
            {
                CurrentDBRevitDocExchanges.Id = value;
                OnPropertyChanged();
            }
        }

        public string RevitDocExchangeType
        {
            get => CurrentDBRevitDocExchanges.RevitDocExchangeType;
            set
            {
                CurrentDBRevitDocExchanges.RevitDocExchangeType = value;
                OnPropertyChanged();
            }
        }

        public string SettingName
        {  
            get => CurrentDBRevitDocExchanges.SettingName;
            set
            {
                CurrentDBRevitDocExchanges.SettingName = value;
                OnPropertyChanged();
            } 
        }
        
        public string SettingResultPath
        {
            get => CurrentDBRevitDocExchanges.SettingResultPath;
            set
            {
                CurrentDBRevitDocExchanges.SettingResultPath = value;
                OnPropertyChanged();
            }
        }

        public int SettingCountItem
        {
            get => CurrentDBRevitDocExchanges.SettingCountItem;
            set
            {
                CurrentDBRevitDocExchanges.SettingCountItem = value;
                OnPropertyChanged();
            }
        }

        public string SettingDBFilePath
        {
            get => CurrentDBRevitDocExchanges.SettingDBFilePath;
            set
            {
                CurrentDBRevitDocExchanges.SettingDBFilePath = value;
                OnPropertyChanged();
            }
        }

        public bool IsSelected
        {
            get => _isSelected; 
            set
            {
                _isSelected = value;
                OnPropertyChanged();
            }
        }

        protected void OnPropertyChanged([CallerMemberName] string name = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
