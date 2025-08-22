using KPLN_Clashes_Ribbon.Services;
using KPLN_Library_SQLiteWorker;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows.Media;

namespace KPLN_Clashes_Ribbon.Core.Reports
{
    /// <summary>
    /// Данные по экземплярам отчетов, объедененные в группы из таблицы групп отчетов
    /// </summary>
    public sealed class Report : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private int _progress = 0;
        private int _delegationProgress = 0;
        private int _approvedProgress = 0;
        private System.Windows.Visibility _pbEnabled = System.Windows.Visibility.Collapsed;
        private System.Windows.Visibility _isGroupEnabled = System.Windows.Visibility.Visible;
        private bool _isReportVisible = true;
        private int _id;
        private int _reportGroupId;
        private ReportGroup _reportGroup;
        private string _name;
        private ClashesMainCollection.KPItemStatus _status;
        private string _path;
        private string _dateCreated;
        private string _userCreated;
        private string _dateLast;
        private string _userLast;
        private Source.Source _source;
        private SolidColorBrush _fill = new SolidColorBrush(Color.FromArgb(255, 255, 255, 255));
        private readonly Services.SQLite.SQLiteService_MainDB _sqliteService_MainDB = new Services.SQLite.SQLiteService_MainDB();
        
        #region Данные из БД
        [Key]
        public int Id
        {
            get => _id;
            set
            {
                _id = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Id группы отчетов
        /// </summary>
        [ForeignKey(nameof(ReportGroup))]
        public int ReportGroupId
        {
            get => _reportGroupId;
            set
            {
                _reportGroupId = value;
                NotifyPropertyChanged();
            }
        }

        public string Name
        {
            get => _name;
            set
            {
                _name = value;
                NotifyPropertyChanged();
            }
        }

        public ClashesMainCollection.KPItemStatus Status
        {
            get => _status;
            set
            {
                _status = value;

                switch (value)
                {
                    case ClashesMainCollection.KPItemStatus.Closed:
                        Source = new Source.Source(ClashesMainCollection.KPIcon.Instance_Closed);
                        break;
                    case ClashesMainCollection.KPItemStatus.Delegated:
                        Source = new Source.Source(ClashesMainCollection.KPIcon.Instance_Delegated);
                        break;
                    case ClashesMainCollection.KPItemStatus.Approved:
                        Source = new Source.Source(ClashesMainCollection.KPIcon.Instance_Approved);
                        break;
                    default:
                        Source = new Source.Source(ClashesMainCollection.KPIcon.Instance);
                        break;
                }

                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Ссылка на БД отчетов
        /// </summary>
        public string PathToReportInstance
        {
            get => _path;
            set
            {
                _path = value;
                NotifyPropertyChanged();
            }
        }

        public string DateCreated
        {
            get => _dateCreated;
            set
            {
                _dateCreated = value;
                NotifyPropertyChanged();
            }
        }

        public string UserCreated
        {
            get => _userCreated;
            set
            {
                _userCreated = value;
                NotifyPropertyChanged();
            }
        }

        public string DateLast
        {
            get => _dateLast;
            set
            {
                _dateLast = value;
                NotifyPropertyChanged();
            }
        }

        public string UserLast
        {
            get => _userLast;
            set
            {
                _userLast = value;
                NotifyPropertyChanged();
            }
        }
        #endregion

        public ReportGroup ReportGroup
        {
            get
            {
                if (_reportGroup == null)
                    _reportGroup = _sqliteService_MainDB.GetReportGroup_ById(ReportGroupId);

                return _reportGroup;
            }
        }

        #region Дополнительная визуализация
        public int Progress
        {
            get => _progress;
            set
            {
                _progress = value;
                NotifyPropertyChanged();
            }
        }

        public int DelegationProgress
        {
            get => _delegationProgress;
            set
            {
                _delegationProgress = value;
                NotifyPropertyChanged();
            }
        }

        public int ApprovedProgress
        {
            get => _approvedProgress;
            set
            {
                _approvedProgress = value;
                NotifyPropertyChanged();
            }
        }

        public SolidColorBrush Fill_Default
        {
            get
            {
                if (IsGroupEnabled == System.Windows.Visibility.Visible)
                    return new SolidColorBrush(Color.FromArgb(225, 255, 255, 255));

                return new SolidColorBrush(Color.FromArgb(255, 155, 155, 155));
            }
        }

        public System.Windows.Visibility IsGroupEnabled
        {
            get => _isGroupEnabled;
            set
            {
                _isGroupEnabled = value;
                NotifyPropertyChanged();
                
                _fill = Fill_Default;
            }
        }

        public bool IsReportVisible
        {
            get => _isReportVisible;
            set
            {
                _isReportVisible = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Вдимость прогресс-бара в instance
        /// </summary>
        public System.Windows.Visibility PbEnabled
        {
            get => _pbEnabled;
            set
            {
                _pbEnabled = value;
                NotifyPropertyChanged();
            }
        }

        public System.Windows.Visibility AdminControllsVisibility
        {
            get
            {
                if (DBMainService.CurrentUserDBSubDepartment.Id == 8)
                    return IsGroupEnabled;

                return System.Windows.Visibility.Collapsed;
            }
        }

        public System.Windows.Visibility IsGroupNotEnabled
        {
            get
            {
                if (_isGroupEnabled == System.Windows.Visibility.Visible)
                    return System.Windows.Visibility.Collapsed;

                return System.Windows.Visibility.Visible;
            }
            set
            {
                _isGroupEnabled = value;
                NotifyPropertyChanged();
                
                _fill = Fill_Default;
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
        #endregion

        public void GetProgress()
        {
            Task t2 = Task.Run(() =>
            {
                UpdateProgress();
            });
        }

        private void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void UpdateProgress()
        {
            int max = 0;
            int done = 0;
            int delegated = 0;
            int approved = 0;
            Services.SQLite.SQLiteService_ReportItemsDB sqliteService_ReportInstanceDB = new Services.SQLite.SQLiteService_ReportItemsDB(PathToReportInstance);
            foreach (ReportItem ri in sqliteService_ReportInstanceDB.GetAllReporItems())
            {
                max++;

                if (ri.Status == ClashesMainCollection.KPItemStatus.Closed)
                    done++;
                else if (ri.Status == ClashesMainCollection.KPItemStatus.Delegated)
                    delegated++;
                else if (ri.Status == ClashesMainCollection.KPItemStatus.Approved)
                    approved++;
            }

            Progress = (int)Math.Round((double)(done * 100 / max));
            DelegationProgress = (int)Math.Round((double)(delegated * 100 / max));
            ApprovedProgress = (int)Math.Round((double)(approved * 100 / max));

            // Устанавливаю статус для смены пиктограммы при условии что все коллизии просмотрены (делегированы, либо устранены, либо открыты)
            if (done == max)
                Status = ClashesMainCollection.KPItemStatus.Closed;
            else if (done + delegated == max && delegated > 0 || (done + delegated + approved == max && delegated > 0 && approved > 0))
                Status = ClashesMainCollection.KPItemStatus.Delegated;
            else if (done + approved == max && approved > 0)
                Status = ClashesMainCollection.KPItemStatus.Approved;
            else
                Status = ClashesMainCollection.KPItemStatus.Opened;
            
            _sqliteService_MainDB.UpdateItemStatus_ByTableAndItemId(Status, MainDB_Enumerator.Reports, Id);

            PbEnabled = System.Windows.Visibility.Visible;
        }
    }
}
