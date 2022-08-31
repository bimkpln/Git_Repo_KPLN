using KPLN_Library_DataBase.Controll;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_Library_DataBase.Collections
{
    public class DbProject : DbElement, INotifyPropertyChanged, IDisposable
    {
        public static ObservableCollection<DbProject> GetAllProjects(ObservableCollection<DbProjectInfo> projectInfo)
        {
            ObservableCollection<DbProject> projects = new ObservableCollection<DbProject>();
            foreach (DbProjectInfo projectData in projectInfo)
            {
                projects.Add(new DbProject(projectData));
            }
            return projects;
        }
        private DbProject(DbProjectInfo projectData)
        {
            _id = projectData.Id;
            _name = projectData.Name;
            _users = projectData.Users;
            _documents = new ObservableCollection<DbDocument>();
            _code = projectData.Code;
            _keys = projectData.Keys;
        }
        public void JoinDocumentsFromList(ObservableCollection<DbDocument> documents)
        {
            foreach (DbDocument doc in documents)
            {
                if (doc.Project != null)
                {
                    if (doc.Project.Id == Id)
                    {
                        _documents.Add(doc);
                    }
                }
            }
        }
        public override string TableName 
        {
            get
            {
                return "Projects";
            }
        }
        private string _name { get; set; }
        public string Name
        {
            get { return _name; }
            set
            {
                if (SQLiteDBUtills.SetValue(this, "Name", value))
                {
                    _name = value;
                    NotifyPropertyChanged();
                }
            }
        }
        private ObservableCollection<DbUser> _users { get; set; }
        public ObservableCollection<DbUser> Users
        {
            get { return _users; }
            set
            {
                List<string> userdata = new List<string>();
                foreach (DbUser user in value)
                {
                    userdata.Add(user.SystemName);
                }
                string _userdata = string.Format("*{0}*", string.Join("*", userdata));
                if (SQLiteDBUtills.SetValue(this, "Users", _userdata))
                {
                    _users = DbControll.GetUsersByNames(userdata);
                    NotifyPropertyChanged();
                }
            }
        }
        private ObservableCollection<DbDocument> _documents { get; set; }
        public ObservableCollection<DbDocument> Documents
        {
            get { return _documents; }
            set
            {
                foreach (DbDocument doc in value)
                {
                    doc.Project = this;
                }
                _documents = DbControll.GetDocumentsByProject(this);
                NotifyPropertyChanged();
            }
        }
        private string _code { get; set; }
        public string Code
        {
            get { return _code; }
            set
            {
                if (SQLiteDBUtills.SetValue(this, "Code", value))
                {
                    _code = value;
                    NotifyPropertyChanged();
                }
            }
        }
        private string _keys { get; set; }
        public string Keys
        {
            get { return _keys; }
            set
            {
                if (SQLiteDBUtills.SetValue(this, "Keys", value))
                {
                    _keys = value;
                    NotifyPropertyChanged();
                }
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;
        public void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
