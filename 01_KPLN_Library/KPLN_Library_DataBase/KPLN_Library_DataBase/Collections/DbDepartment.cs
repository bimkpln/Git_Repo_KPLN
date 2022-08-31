using KPLN_Library_DataBase.Controll;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_Library_DataBase.Collections
{
    public class DbDepartment : DbElement, INotifyPropertyChanged, IDisposable
    {
        public static ObservableCollection<DbDepartment> GetAllDepartments(ObservableCollection<DbDepartmentInfo> departmentInfo)
        {
            ObservableCollection<DbDepartment> departments = new ObservableCollection<DbDepartment>();
            foreach (DbDepartmentInfo departmentData in departmentInfo)
            {
                departments.Add(new DbDepartment(departmentData));
            }
            return departments;
        }
        private DbDepartment(DbDepartmentInfo departmentData)
        {
            _id = departmentData.Id;
            _name = departmentData.Name;
            _code = departmentData.Code;
        }
        public override string TableName
        {
            get
            {
                return "Departments";
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
        public event PropertyChangedEventHandler PropertyChanged;
        public void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
