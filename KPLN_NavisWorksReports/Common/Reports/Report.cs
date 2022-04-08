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

namespace KPLN_NavisWorksReports.Common.Reports
{
    public class Report : INotifyPropertyChanged
    {
        private int _Progress { get; set; } = 0;
        public int Progress
        {
            get
            {
                return _Progress;
            }
            set
            {
                _Progress = value;
                NotifyPropertyChanged();
            }
        }
        public void UpdateProgress()
        {
            int max = 0;
            int done = 0;
            foreach (ReportInstance ri in ReportInstance.GetReportInstances(Path))
            {
                max++;
                if (ri.Status != Collections.Status.Opened)
                {
                    done++;
                }
            }
            int result = (int)Math.Round((double)(done * 100 / max));
            Progress = result;
            PbEnabled = System.Windows.Visibility.Visible;
        }
        private System.Windows.Visibility _PbEnabled { get; set; } = System.Windows.Visibility.Collapsed;
        private System.Windows.Visibility _IsGroupEnabled { get; set; } = System.Windows.Visibility.Visible;
        private int _Id { get; set; }
        private int _GroupId { get; set; }
        private string _Name { get; set; }
        private int _Status { get; set; }
        private string _Path { get; set; }
        private string _DateCreated { get; set; }
        private string _UserCreated { get; set; }
        private string _DateLast { get; set; }
        private string _UserLast { get; set; }
        private Source.Source _Source { get; set; }
        public SolidColorBrush _Fill_Default
        {
            get
            {
                if (IsGroupEnabled == System.Windows.Visibility.Visible)
                { return new SolidColorBrush(System.Windows.Media.Color.FromArgb(225, 255, 255, 255)); }
                return new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 155, 155, 155));
            }
        }
        private SolidColorBrush _Fill { get; set; } = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 255, 255));
        public System.Windows.Visibility IsGroupEnabled
        {
            get
            {
                return _IsGroupEnabled;
            }
            set
            {
                _IsGroupEnabled = value;
                NotifyPropertyChanged();
                _Fill = _Fill_Default;
            }
        }
        public System.Windows.Visibility PbEnabled
        {
            get
            {
                return _PbEnabled;
            }
            set
            {
                _PbEnabled = value;
                NotifyPropertyChanged();
            }
        }
        public System.Windows.Visibility AdminControllsVisibility
        {
            get
            {
                if (KPLN_Loader.Preferences.User.Department.Id != 4)
                {
                    return System.Windows.Visibility.Collapsed;
                }
                return IsGroupEnabled;
            }
        }
        public System.Windows.Visibility IsGroupNotEnabled
        {
            get
            {
                if (_IsGroupEnabled == System.Windows.Visibility.Visible)
                { return System.Windows.Visibility.Collapsed; } 
                return System.Windows.Visibility.Visible;
            }
            set
            {
                _IsGroupEnabled = value;
                NotifyPropertyChanged();
                _Fill = _Fill_Default;
            }
        }
        public SolidColorBrush Fill
        {
            get
            {
                return _Fill;
            }
            set
            {
                _Fill = value;
                NotifyPropertyChanged();
            }
        }
        public int Id
        {
            get
            {
                return _Id;
            }
            set
            {
                _Id = value;
                NotifyPropertyChanged();
            }
        }
        public int GroupId
        {
            get
            {
                return _GroupId;
            }
            set
            {
                _GroupId = value;
                NotifyPropertyChanged();
            }
        }
        public string Name
        {
            get
            {
                return _Name;
            }
            set
            {
                _Name = value;
                NotifyPropertyChanged();
            }
        }
        public int Status
        {
            get
            {
                return _Status;
            }
            set
            {
                if (value != 1)
                {
                    Source = new Source.Source(Collections.Icon.Instance);
                }
                else
                { 
                    Source = new Source.Source(Collections.Icon.Instance_Closed);
                }
                _Status = value;
                NotifyPropertyChanged();
            }
        }
        public string Path
        {
            get
            {
                return _Path;
            }
            set
            {
                _Path = value;
                NotifyPropertyChanged();
            }
        }
        public string DateCreated
        {
            get
            {
                return _DateCreated;
            }
            set
            {
                _DateCreated = value;
                NotifyPropertyChanged();
            }
        }
        public string UserCreated
        {
            get
            {
                return _UserCreated;
            }
            set
            {
                _UserCreated = value;
                NotifyPropertyChanged();
            }
        }
        public string DateLast
        {
            get
            {
                return _DateLast;
            }
            set
            {
                _DateLast = value;
                NotifyPropertyChanged();
            }
        }
        public string UserLast
        {
            get
            {
                return _UserLast;
            }
            set
            {
                _UserLast = value;
                NotifyPropertyChanged();
            }
        }
        public Source.Source Source
        {
            get
            {
                return _Source;
            }
            set
            {
                _Source = value;
                NotifyPropertyChanged();
            }
        }
        public void GetProgress()
        {
            Task t2 = Task.Run(() =>
            { UpdateProgress(); });
        }
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
        public static ObservableCollection<Report> GetReports()
        {
            ObservableCollection<Report> reports = new ObservableCollection<Report>();
            try
            {
                SQLiteConnection db = new SQLiteConnection(string.Format(@"Data Source=Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_NwcReports.db;Version=3;"));
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
        public event PropertyChangedEventHandler PropertyChanged;
        public void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
