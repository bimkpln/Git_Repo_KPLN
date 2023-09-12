using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Runtime.CompilerServices;
using System.Windows.Media;
using static KPLN_Clashes_Ribbon.Core.ClashesMainCollection;

namespace KPLN_Clashes_Ribbon.Core.Reports
{
    /// <summary>
    /// Общая группа отчетов из таблицы групп отчетов
    /// </summary>
    public sealed class ReportGroup : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private ObservableCollection<Report> _reports = new ObservableCollection<Report>();
        private int _id;
        private int _projectId;
        private string _name;
        private KPItemStatus _status;
        private string _dateCreated;
        private string _userCreated;
        private string _dateLast;
        private string _userLast;
        private Source.Source _source;
        private SolidColorBrush _fill;
        private bool _isEnabled = true;
        private bool _isExpandedItem = false;
        private static ProjectDbService _libProjectDbService;

        #region Данные из БД
        [Key]
        public int Id
        {
            get => _id;
            set
            {
                if (_id != value)
                {
                    _id = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public int ProjectId
        {
            get => _projectId;
            set
            {
                if (_projectId != value)
                {
                    _projectId = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public string Name
        {
            get => _name;
            set
            {
                if (_name != value)
                {
                    _name = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public KPItemStatus Status
        {
            get => _status;
            set
            {
                UpdatePropertiesBasedOnStatus();
                _status = value;
                NotifyPropertyChanged();
            }
        }

        public string DateCreated
        {
            get => _dateCreated;
            set
            {
                if (_dateCreated != value)
                {
                    _dateCreated = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public string UserCreated
        {
            get => _userCreated;
            set
            {
                if (_userCreated != value)
                {
                    _userCreated = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public string DateLast
        {
            get => _dateLast;
            set
            {
                if (_dateLast != value)
                {
                    _dateLast = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public string UserLast
        {
            get => _userLast;
            set
            {
                if (_userLast != value)
                {
                    _userLast = value;
                    NotifyPropertyChanged();
                    NotifySelectedPropertyChanged(nameof(LastUserFullName));
                }
            }
        }
        #endregion

        #region Дополнительная визуализация
        /// <summary>
        /// Специальный формат для вывода в окно пользователя
        /// </summary>
        public string LastUserFullName
        {
            get => $"{KPLN_Loader.Application.CurrentRevitUser.Name} {KPLN_Loader.Application.CurrentRevitUser.Surname}";
        }

        /// <summary>
        /// Специальный формат для вывода в окно пользователя
        /// </summary>
        public string UserCreatedFullName
        {
            get => $"{KPLN_Loader.Application.CurrentRevitUser.Name} {KPLN_Loader.Application.CurrentRevitUser.Surname}";
        }

        /// <summary>
        /// Специальный формат для вывода в окно пользователя
        /// </summary>
        public string GroupName
        {
            get
            {
                DBProject dBProject = LibProjectDbService.GetDBProject_ByProjectId(ProjectId);
                return $"[{dBProject.Code}]: {Name} ({dBProject.Name})";
            }
        }

        public bool IsEnabled
        {
            get => _isEnabled;
            set
            {
                if (_isEnabled != value)
                {
                    _isEnabled = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public SolidColorBrush Fill
        {
            get => _fill;
            set
            {
                if (_fill != value)
                {
                    _fill = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public ObservableCollection<Report> Reports
        {
            get => _reports;
            set
            {
                if (_reports != value)
                {
                    _reports = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public bool IsExpandedItem
        {
            get => _isExpandedItem;
            set
            {
                if (_isExpandedItem != value)
                {
                    _isExpandedItem = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public System.Windows.Visibility AdminControllsVisibility
        {
            get
            {
                if (KPLN_Loader.Application.CurrentRevitUser.SubDepartmentId == 8)
                {
                    return System.Windows.Visibility.Visible;
                }

                return System.Windows.Visibility.Collapsed;
            }
        }

        public System.Windows.Visibility AdminControllsVisibilityAdd
        {
            get
            {
                if (IsEnabled)
                    return System.Windows.Visibility.Visible;
                else
                    return System.Windows.Visibility.Collapsed;
            }
        }

        public Source.Source Source
        {
            get => _source;
            set
            {
                if (_source != value)
                {
                    _source = value;
                    NotifyPropertyChanged();
                }
            }
        }
        #endregion

        private void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void NotifySelectedPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        /// <summary>
        /// Синглтон сервис из библиотеки KPLN_Library_SQLiteWorker
        /// </summary>
        private static ProjectDbService LibProjectDbService
        {
            get
            {
                if (_libProjectDbService == null)
                {
                    CreatorProjectDbService creatorPrjDbService = new CreatorProjectDbService();
                    _libProjectDbService = (ProjectDbService)creatorPrjDbService.CreateService();
                }

                return _libProjectDbService;
            }
        }

        /// <summary>
        /// Обновление свойств отюъекта, в зависимости от стаутса ReportGroup
        /// </summary>
        private void UpdatePropertiesBasedOnStatus()
        {
            switch (_status)
            {
                case KPItemStatus.New:
                    Source = new Source.Source(KPIcon.Report_New);
                    IsExpandedItem = true;
                    Fill = new SolidColorBrush(Color.FromArgb(255, 86, 88, 211));
                    IsEnabled = true;
                    break;
                case KPItemStatus.Opened:
                    Source = new Source.Source(KPIcon.Report);
                    IsExpandedItem = true;
                    Fill = new SolidColorBrush(Color.FromArgb(255, 86, 156, 211));
                    IsEnabled = true;
                    break;
                case KPItemStatus.Closed:
                    Source = new Source.Source(KPIcon.Report_Closed);
                    IsExpandedItem = false;
                    Fill = new SolidColorBrush(Color.FromArgb(255, 78, 97, 112));
                    IsEnabled = false;
                    break;
            }
        }
    }
}
