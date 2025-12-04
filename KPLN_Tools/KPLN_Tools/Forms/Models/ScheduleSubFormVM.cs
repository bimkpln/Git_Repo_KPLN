using KPLN_Tools.Forms.Models.Core;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;

namespace KPLN_Tools.Forms.Models
{
    public sealed class ScheduleSubFormVM : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        
        private string _prefix;
        private string _startNumber;
        private bool _autoIncrement = true;

        public ScheduleSubFormVM(string header, string startNumber)
        {
            Header = $"Столбец: \"{header}\"";
            StartNumber = startNumber ?? string.Empty;

            RunCmd = new RelayCommand<object>(Run);
            CloseWindowCmd = new RelayCommand<object>(CloseWindow);
        }

        /// <summary>
        /// Заголовок окна
        /// </summary>
        public string Header { get; }

        /// <summary>
        /// Приставка
        /// </summary>
        public string Prefix
        {
            get => _prefix;
            set
            {
                if (_prefix == value)
                    return;

                _prefix = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Номер
        /// </summary>
        public string StartNumber
        {
            get => _startNumber;
            set
            {
                if (_startNumber == value) 
                    return;
                
                _startNumber = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Сделать автоинкремент?
        /// </summary>
        public bool AutoIncrement
        {
            get => _autoIncrement;
            set
            {
                if (_autoIncrement == value)
                    return;

                _autoIncrement = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Итоговое значение, которое далее используется для ячеек
        /// </summary>
        public string FullCellData => $"{Prefix}{StartNumber}";

        /// <summary>
        /// Команда: Запуск
        /// </summary>
        public ICommand RunCmd { get; }

        /// <summary>
        /// Команда: Закрыть окно
        /// </summary>
        public ICommand CloseWindowCmd { get; }

        public void Run(object windObj)
        {
            if (windObj is Window window)
            {
                window.DialogResult = true;
                window.Close();
            }
        }

        public void CloseWindow(object windObj)
        {
            if (windObj is Window window)
            {
                window.DialogResult = false;
                window.Close();
            }
        }

        private void OnPropertyChanged([CallerMemberName] string propName = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propName));
    }
}
