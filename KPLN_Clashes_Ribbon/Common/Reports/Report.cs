using KPLN_Clashes_Ribbon.Tools;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using static KPLN_Loader.Output.Output;

namespace KPLN_Clashes_Ribbon.Common.Reports
{
    /// <summary>
    /// Данные по экземплярам отчетов, объедененные в группы из таблицы групп отчетов
    /// </summary>
    public sealed class Report : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        
        private int _progress = 0;

        private int _delegationProgress = 0;

        private System.Windows.Visibility _pbEnabled = System.Windows.Visibility.Collapsed;

        private System.Windows.Visibility _isGroupEnabled = System.Windows.Visibility.Visible;

        private int _id;

        private int _groupId;
        
        private string _name;
        
        private int _status;
        
        private string _path;

        private string _dateCreated;

        private string _userCreated;

        private string _dateLast;

        private string _userLast;

        private Source.Source _source;

        private SolidColorBrush _fill = new SolidColorBrush(Color.FromArgb(255, 255, 255, 255));

        private Report(int id, int groupId, string name, int status, string path, string dateCreated, string userCreated, string dateLast, string userLast)
        {
            Id = id;
            GroupId = groupId;
            Name = name;
            Status = status;
            Path = path;
            DateCreated = dateCreated;
            UserCreated = userCreated;
            DateLast = dateLast;
            UserLast = userLast;
        }
        
        public int Progress
        {
            get
            {
                return _progress;
            }
            set
            {
                _progress = value;
                NotifyPropertyChanged();
            }
        }

        public int DelegationProgress
        {
            get
            {
                return _delegationProgress;
            }
            set 
            {
                _delegationProgress = value; 
                NotifyPropertyChanged(); 
            }
        }

        public SolidColorBrush _Fill_Default
        {
            get
            {
                if (IsGroupEnabled == System.Windows.Visibility.Visible)
                { return new SolidColorBrush(Color.FromArgb(225, 255, 255, 255)); }
                return new SolidColorBrush(Color.FromArgb(255, 155, 155, 155));
            }
        }

        public System.Windows.Visibility IsGroupEnabled
        {
            get
            {
                return _isGroupEnabled;
            }
            set
            {
                _isGroupEnabled = value;
                NotifyPropertyChanged();
                _fill = _Fill_Default;
            }
        }

        public System.Windows.Visibility PbEnabled
        {
            get
            {
                return _pbEnabled;
            }
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
                if (KPLN_Loader.Preferences.User.Department.Id == 4 || KPLN_Loader.Preferences.User.Department.Id == 6)
                {
                    return IsGroupEnabled;
                }

                return System.Windows.Visibility.Collapsed;
            }
        }

        public System.Windows.Visibility IsGroupNotEnabled
        {
            get
            {
                if (_isGroupEnabled == System.Windows.Visibility.Visible)
                { return System.Windows.Visibility.Collapsed; }
                return System.Windows.Visibility.Visible;
            }
            set
            {
                _isGroupEnabled = value;
                NotifyPropertyChanged();
                _fill = _Fill_Default;
            }
        }

        public int Id
        {
            get
            {
                return _id;
            }
            set
            {
                _id = value;
                NotifyPropertyChanged();
            }
        }

        public int GroupId
        {
            get
            {
                return _groupId;
            }
            set
            {
                _groupId = value;
                NotifyPropertyChanged();
            }
        }

        public string Name
        {
            get
            {
                return _name;
            }
            set
            {
                _name = value;
                NotifyPropertyChanged();
            }
        }

        public int Status
        {
            get
            {
                return _status;
            }
            set
            {
                if (value == 1)
                { Source = new Source.Source(Collections.Icon.Instance_Closed); }
                else if (value == 2)
                { Source = new Source.Source(Collections.Icon.Instance_Delegated); }
                else 
                { Source = new Source.Source(Collections.Icon.Instance); }

                _status = value;

                NotifyPropertyChanged();
            }
        }

        public string Path
        {
            get
            {
                return _path;
            }
            set
            {
                _path = value;
                NotifyPropertyChanged();
            }
        }

        public string DateCreated
        {
            get
            {
                return _dateCreated;
            }
            set
            {
                _dateCreated = value;
                NotifyPropertyChanged();
            }
        }

        public string UserCreated
        {
            get
            {
                return _userCreated;
            }
            set
            {
                _userCreated = value;
                NotifyPropertyChanged();
            }
        }

        public string DateLast
        {
            get
            {
                return _dateLast;
            }
            set
            {
                _dateLast = value;
                NotifyPropertyChanged();
            }
        }

        public string UserLast
        {
            get
            {
                return _userLast;
            }
            set
            {
                _userLast = value;
                NotifyPropertyChanged();
            }
        }

        public Source.Source Source
        {
            get
            {
                return _source;
            }
            set
            {
                _source = value;
                NotifyPropertyChanged();
            }
        }

        public SolidColorBrush Fill
        {
            get
            {
                return _fill;
            }
            set
            {
                _fill = value;
                NotifyPropertyChanged();
            }
        }

        public static ObservableCollection<Report> GetReports()
        {
            ObservableCollection<Report> reports = new ObservableCollection<Report>();
            try
            {
                SQLiteConnection db = new SQLiteConnection(
                    string.Format(
                        @"Data Source=Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_NwcReports.db;Version=3;")
                    );

                try
                {
                    db.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand("SELECT * FROM Reports", db))
                    {
                        using (SQLiteDataReader rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                if (rdr.GetInt32(0) == -1) { continue; }
                                try
                                {
                                    reports.Add(new Report(rdr.GetInt32(0),
                                        rdr.GetInt32(1),
                                        rdr.GetString(2),
                                        rdr.GetInt32(3),
                                        rdr.GetString(4),
                                        rdr.GetString(5),
                                        rdr.GetString(6),
                                        rdr.GetString(7),
                                        rdr.GetString(8)));
                                }
                                catch (Exception e)
                                {
                                    PrintError(e);
                                }

                            }
                        }
                    }
                    db.Close();
                }
                catch (Exception)
                {
                    db.Close();
                }
            }
            catch (Exception) { }
            return reports;
        }

        public void UpdateProgress()
        {
            int max = 0;
            int done = 0;
            int delegated = 0;
            foreach (ReportInstance ri in ReportInstance.GetReportInstances(Path))
            {
                max++;
                
                if (ri.Status != Collections.Status.Opened && ri.Status != Collections.Status.Delegated)
                { done++; }
                else if (ri.Status == Collections.Status.Delegated)
                { delegated++; }
            }
            
            int doneCount = (int)Math.Round((double)(done * 100 / max));
            Progress = doneCount;

            int delegatedCount = (int)Math.Round((double)(delegated * 100 / max));
            DelegationProgress = delegatedCount;

            // Устанавливаю статус для смены пиктограммы при условии что все коллизии просмотрены (делегированы, либо устранены)
            if (done + delegated == max && delegated > 0)
            {
                Status = 2;
                DbController.SetInstanceValue(Path, Id, "STATUS", 2);
            }

            PbEnabled = System.Windows.Visibility.Visible;
        }

        public void GetProgress()
        {
            Task t2 = Task.Run(() =>
            { UpdateProgress(); });
        }

        public void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
