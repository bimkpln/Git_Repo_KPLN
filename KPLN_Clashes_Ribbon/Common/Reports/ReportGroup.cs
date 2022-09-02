using KPLN_Library_DataBase.Collections;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Media;
using static KPLN_Loader.Output.Output;
using static KPLN_Clashes_Ribbon.Common.Collections;

namespace KPLN_Clashes_Ribbon.Common.Reports
{
    public class ReportGroup : INotifyPropertyChanged
    {
        private ObservableCollection<Report> _Reports { get; set; } = new ObservableCollection<Report>();
        
        private int _Id { get; set; }
        
        private int _ProjectId { get; set; }
        
        private string _Name { get; set; }
        
        private int _Status { get; set; }
        
        private string _DateCreated { get; set; }
        
        private string _UserCreated { get; set; }
        
        private string _DateLast { get; set; }
        
        private string _UserLast { get; set; }
        
        private Source.Source _Source { get; set; }
        
        private SolidColorBrush _Fill { get; set; }
        
        private bool _IsEnabled { get; set; } = true;
        
        public string DBUserLast { get; set; }
        
        public string DBUserCreated { get; set; }
        
        public bool IsEnabled
        {
            get
            {
                return _IsEnabled;
            }
            set
            {
                _IsEnabled = value;
                NotifyPropertyChanged();
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
        public ObservableCollection<Report> Reports
        {
            get
            {
                return _Reports;
            }
            set
            {
                _Reports = value;
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
        public int ProjectId
        {
            get
            {
                return _ProjectId;
            }
            set
            {
                _ProjectId = value;
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
        private bool _IsExpandedItem { get; set; } = false;
        public bool IsExpandedItem
        {
            get
            {
                return _IsExpandedItem;
            }
            set
            {
                _IsExpandedItem = value;
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
                _Status = value;
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
                return System.Windows.Visibility.Visible;
            }
        }
        public System.Windows.Visibility AdminControllsVisibilityAdd
        {
            get
            {
                if (KPLN_Loader.Preferences.User.Department.Id != 4)
                {
                    
                }
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
                DBUserCreated = ModuleData.GetUserBySystemName(value);
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
                DBUserLast = ModuleData.GetUserBySystemName(value);
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
                SQLiteConnection db = new SQLiteConnection(string.Format(@"Data Source=Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_NwcReports.db;Version=3;"));
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
        public event PropertyChangedEventHandler PropertyChanged;
        public void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
