using KPLN_Library_DataBase.Controll;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_Library_DataBase.Collections
{
    public class DbSubDepartment : DbElement, INotifyPropertyChanged, IDisposable
    {
        public static ObservableCollection<DbSubDepartment> GetAllSubDepartments(ObservableCollection<DbSubDepartmentInfo> subDepartmentInfo)
        {
            ObservableCollection<DbSubDepartment> subDepartments = new ObservableCollection<DbSubDepartment>();
            foreach (DbSubDepartmentInfo subDepartmentData in subDepartmentInfo)
            {
                subDepartments.Add(new DbSubDepartment(subDepartmentData));
            }
            return subDepartments;
        }
        private DbSubDepartment(DbSubDepartmentInfo subDepartmentData)
        {
            _id = subDepartmentData.Id;
            _name = subDepartmentData.Name;
            _code = subDepartmentData.Code;
            _codeus = subDepartmentData.CodeUS;
        }
        public override string TableName
        {
            get
            {
                return "SubDepartments";
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
        private string _codeus { get; set; }
        public string CodeUS
        {
            get { return _codeus; }
            set
            {
                if (SQLiteDBUtills.SetValue(this, "CodeUS", value))
                {
                    _codeus = value;
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
