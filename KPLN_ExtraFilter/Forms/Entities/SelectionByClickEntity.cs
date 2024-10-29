using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_ExtraFilter.Entities
{
    public class SelectionByClickEntity : INotifyPropertyChanged
    {
        private bool _sameCategory;
        private bool _sameFamily;
        private bool _sameType;
        private bool _sameWorkset;
        private bool _model;
        private bool _currentView;
        private bool _belongGroup;

        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Одинаковой категории
        /// </summary>
        public bool What_SameCategory
        {
            get => _sameCategory;
            set
            {
                if (_sameCategory != value)
                {
                    _sameCategory = value;
                    NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Одинакового семейства
        /// </summary>
        public bool What_SameFamily
        {
            get => _sameFamily;
            set
            {
                if (_sameFamily != value)
                {
                    _sameFamily = value;
                    NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Одинакового типа
        /// </summary>
        public bool What_SameType
        {
            get => _sameType;
            set
            {
                if (_sameType != value)
                {
                    _sameType = value;
                    NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Одного рабочего набора
        /// </summary>
        public bool What_Workset
        {
            get => _sameWorkset;
            set
            {
                if (_sameWorkset != value)
                {
                    _sameWorkset = value;
                    NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// В модели
        /// </summary>
        public bool Where_Model
        {
            get => _model;
            set
            {
                if (_model != value)
                {
                    _model = value;
                    NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// На виде
        /// </summary>
        public bool Where_CurrentView
        {
            get => _currentView;
            set
            {
                if (_currentView != value)
                {
                    _currentView = value;
                    NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Исключить элементы групп
        /// </summary>
        public bool Belong_Group
        {
            get => _belongGroup;
            set
            {
                if (_belongGroup != value)
                {
                    _belongGroup = value;
                    NotifyPropertyChanged();
                }
            }
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
