using KPLN_Library_ConfigWorker.Core;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_Tools.Forms.Models
{
    public sealed class AutoSaveM : INotifyPropertyChanged, IJsonSerializable
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private bool _isAutoSaveEnabled = true;
        private bool _showWarningWindow = true;
        private int _selectedInterval = 30;
        private bool _isSaveToCopiesEnabled = true;
        private int _copiesCountSelected = 2;

        [JsonConstructor]
        public AutoSaveM(){ }

        /// <summary>
        /// Включить автосохранение.
        /// </summary>
        public bool IsAutoSaveEnabled
        {
            get => _isAutoSaveEnabled;
            set
            {
                if (_isAutoSaveEnabled == value)
                    return;

                _isAutoSaveEnabled = value;
                OnPropertyChanged();

                // Если автосохранение выключено — блокируем 
                OnPropertyChanged(nameof(IntervalOptionsEnabled));
            }
        }

        /// <summary>
        /// Показывать окно «Предупреждение об автосохранении».
        /// </summary>
        public bool ShowWarningWindow
        {
            get => _showWarningWindow;
            set
            {
                if (_showWarningWindow != value)
                {
                    _showWarningWindow = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Выбранный интервал автосохранения (минуты).
        /// </summary>
        public int SelectedInterval

        {
            get => _selectedInterval;
            set
            {
                if (_selectedInterval != value)
                {
                    _selectedInterval = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Доступность выбора интервалов
        /// </summary>
        public bool IntervalOptionsEnabled => _isAutoSaveEnabled;

        /// <summary>
        /// Допустимые интервалы автосохранения.
        /// </summary>
        public IReadOnlyList<int> Intervals { get; } = new List<int> { 2, 3, 5, 10, 15, 20, 30, 60 };

        /// <summary>
        /// Включить автосохранение в отдельные копии.
        /// </summary>
        public bool IsSaveToCopiesEnabled
        {
            get => _isSaveToCopiesEnabled;
            set
            {
                if (_isSaveToCopiesEnabled == value)
                    return;

                _isSaveToCopiesEnabled = value;
                OnPropertyChanged();

                // Если автосохранение выключено — блокируем 
                OnPropertyChanged(nameof(IntervalCopiesCountEnabled));
            }
        }

        /// <summary>
        /// Выбранный колличество копий
        /// </summary>
        public int CopiesCountSelected

        {
            get => _copiesCountSelected;
            set
            {
                if (_copiesCountSelected != value)
                {
                    _copiesCountSelected = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Доступность выбора интервалов количества резервных копий
        /// </summary>
        public bool IntervalCopiesCountEnabled => _isAutoSaveEnabled && _isSaveToCopiesEnabled;

        /// <summary>
        /// Допустимые опции количества резервных копий
        /// </summary>
        public IReadOnlyList<int> CopiesCountVar { get; } = new List<int> { 2, 3, 4, 5 };


        public object ToJson() => new
        {
            this.IsAutoSaveEnabled,
            this.ShowWarningWindow,
            this.SelectedInterval,
            this.IsSaveToCopiesEnabled,
            this.CopiesCountSelected,

        };

        private void OnPropertyChanged([CallerMemberName] string propName = null)
           => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propName));
    }
}
