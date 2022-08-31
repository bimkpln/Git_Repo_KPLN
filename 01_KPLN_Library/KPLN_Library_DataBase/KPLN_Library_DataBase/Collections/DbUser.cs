using KPLN_Library_DataBase.Controll;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_Library_DataBase.Collections
{
    public class DbUser : DbElement, INotifyPropertyChanged, IDisposable
    {
        public static ObservableCollection<DbUser> GetAllUsers(ObservableCollection<DbUserInfo> userInfo)
        {
            ObservableCollection<DbUser> users = new ObservableCollection<DbUser>();
            foreach (DbUserInfo userData in userInfo)
            {
                users.Add(new DbUser(userData));
            }
            return users;
        }
        private DbUser(DbUserInfo userData)
        {
            _systemname = userData.SystemName;
            _name = userData.Name;
            _family = userData.Family;
            _surname = userData.Surname;
            _department = userData.Department;
            _status = userData.Status;
            _connection = userData.Connection;
            _sex = userData.Sex;
        }
        public override string TableName
        {
            get
            {
                return "Users";
            }
        }
        private string _systemname { get; set; }
        public string SystemName
        {
            get { return _systemname; }
            set
            {
                if (SQLiteDBUtills.SetValue(this, "SystemName", value))
                {
                    _systemname = value;
                    NotifyPropertyChanged();
                }
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
        private string _family { get; set; }
        public string Family
        {
            get { return _family; }
            set
            {
                if (SQLiteDBUtills.SetValue(this, "Family", value))
                {
                    _family = value;
                    NotifyPropertyChanged();
                }
            }
        }
        private string _surname { get; set; }
        public string Surname
        {
            get { return _surname; }
            set
            {
                if (SQLiteDBUtills.SetValue(this, "Surname", value))
                {
                    _surname = value;
                    NotifyPropertyChanged();
                }
            }
        }
        private DbDepartment _department { get; set; }
        public DbDepartment Department
        {
            get { return _department; }
            set
            {
                if (SQLiteDBUtills.SetValue(this, "Department", value.Id))
                {
                    _department = value;
                    NotifyPropertyChanged();
                }
            }
        }
        private string _status { get; set; }
        public string Status
        {
            get { return _status; }
            set
            {
                if (SQLiteDBUtills.SetValue(this, "Status", value))
                {
                    _status = value;
                    NotifyPropertyChanged();
                }
            }
        }
        private string _connection { get; set; }
        public string Connection
        {
            get { return _connection; }
            set
            {
                if (SQLiteDBUtills.SetValue(this, "Connection", value))
                {
                    _connection = value;
                    NotifyPropertyChanged();
                }
            }
        }
        private string _sex { get; set; }
        public string Sex
        {
            get { return _sex; }
            set
            {
                if (SQLiteDBUtills.SetValue(this, "Sex", value))
                {
                    _sex = value;
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
