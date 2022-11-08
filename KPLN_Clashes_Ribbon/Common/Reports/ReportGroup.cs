using KPLN_Library_DataBase.Collections;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data.SQLite;
using System.Runtime.CompilerServices;
using System.Windows.Media;
using static KPLN_Clashes_Ribbon.Common.Collections;
using static KPLN_Loader.Output.Output;

namespace KPLN_Clashes_Ribbon.Common.Reports
{
    /// <summary>
    /// Общая группа отчетов из таблицы групп отчетов
    /// </summary>
    internal sealed class ReportGroup : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private ObservableCollection<Report> _reports = new ObservableCollection<Report>();

        private int _id;

        private int _projectId;

        private string _name;

        private int _status;

        private string _dateCreated;

        private string _userCreated;

        private string _dateLast;

        private string _userLast;

        private Source.Source _source;

        private SolidColorBrush _fill;

        private bool _isEnabled = true;

        private bool _isExpandedItem = false;

        private ReportGroup(ObservableCollection<DbProject> dbProjects, int id, int projectId, string name, int status, string dateCreated, string userCreated, string dateLast, string userLast)
        {
            Id = id;
            ProjectId = projectId;
            Name = name;
            foreach (DbProject proj in dbProjects)
            {
                if (projectId == proj.Id)
                {
                    Name = string.Format("[{0}]: {1} ({2})", proj.Code, name, proj.Name);
                }
            }
            Status = status;
            DateCreated = dateCreated;
            UserCreated = userCreated;
            DateLast = dateLast;
            UserLast = userLast;
        }

        public string DBUserLast { get; set; }

        public string DBUserCreated { get; set; }

        public bool IsEnabled
        {
            get
            {
                return _isEnabled;
            }
            set
            {
                _isEnabled = value;
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

        public ObservableCollection<Report> Reports
        {
            get
            {
                return _reports;
            }
            set
            {
                _reports = value;
                NotifyPropertyChanged();
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

        public int ProjectId
        {
            get
            {
                return _projectId;
            }
            set
            {
                _projectId = value;
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

        public bool IsExpandedItem
        {
            get
            {
                return _isExpandedItem;
            }
            set
            {
                _isExpandedItem = value;
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
                if (value == -1)
                {
                    Source = new Source.Source(Icon.Report_New);
                    IsExpandedItem = true;
                    Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 86, 156, 211));
                    IsEnabled = true;
                }
                if (value == 0)
                {
                    Source = new Source.Source(Icon.Report);
                    IsExpandedItem = true;
                    Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 86, 156, 211));
                    IsEnabled = true;
                }
                if (value == 1)
                {
                    Source = new Source.Source(Icon.Report_Closed);
                    IsExpandedItem = false;
                    Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 78, 97, 112));
                    IsEnabled = false;
                }
                _status = value;
                NotifyPropertyChanged();
            }
        }

        public System.Windows.Visibility AdminControllsVisibility
        {
            get
            {
                if (KPLN_Loader.Preferences.User.Department.Id == 4 || KPLN_Loader.Preferences.User.Department.Id == 6)
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
                {
                    return System.Windows.Visibility.Visible;
                }
                else
                {
                    return System.Windows.Visibility.Collapsed;
                }

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
                DBUserCreated = ModuleData.GetUserBySystemName(value);
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
                DBUserLast = ModuleData.GetUserBySystemName(value);
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


        public static ObservableCollection<ReportGroup> GetReportGroups(DbProject project)
        {
            ObservableCollection<ReportGroup> groups = new ObservableCollection<ReportGroup>();
            ObservableCollection<DbProject> dbProjects;
            string sqlCommand;
            if (project != null)
            {
                dbProjects = new ObservableCollection<DbProject>() { project };
                sqlCommand = $"SELECT * FROM ReportGroups Where ProjectId = {project.Id}";
            }
            else
            {
                KPLN_Library_DataBase.DbControll.Update();
                dbProjects = KPLN_Library_DataBase.DbControll.Projects;
                sqlCommand = $"SELECT * FROM ReportGroups";
            }

            try
            {
                SQLiteConnection db = new SQLiteConnection(
                    string.Format(
                        @"Data Source=Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_NwcReports.db;Version=3;")
                    );
                
                try
                {
                    db.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sqlCommand, db))
                    {
                        using (SQLiteDataReader rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                if (rdr.GetInt32(0) == -1) { continue; }
                                try
                                {
                                    groups.Add(new ReportGroup(
                                        dbProjects,
                                        rdr.GetInt32(0),
                                        rdr.GetInt32(1),
                                        rdr.GetString(2),
                                        rdr.GetInt32(3),
                                        rdr.GetString(4),
                                        rdr.GetString(5),
                                        rdr.GetString(6),
                                        rdr.GetString(7)));
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
            return groups;
        }

        public void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

    }
}
