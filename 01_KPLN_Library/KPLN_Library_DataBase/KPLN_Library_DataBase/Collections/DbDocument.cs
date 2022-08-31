using KPLN_DataBase.Controll;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_DataBase.Collections
{
    public class DbDocument : DbElement, INotifyPropertyChanged, IDisposable
    {
        public static ObservableCollection<DbDocument> GetAllDocuments(ObservableCollection<DbDocumentInfo> documentInfo)
        {
            ObservableCollection<DbDocument> documents = new ObservableCollection<DbDocument>();
            foreach (DbDocumentInfo documentData in documentInfo)
            {
                documents.Add(new DbDocument(documentData));
            }
            return documents;
        }
        private DbDocument(DbDocumentInfo documentData)
        {
            _id = documentData.Id;
            _path = documentData.Path;
            _name = documentData.Name;
            _department = documentData.Department;
            _project = documentData.Project;
            _code = documentData.Code;
        }
        public override string TableName
        {
            get
            {
                return "Documents";
            }
        }
        private string _path { get; set; }
        public string Path
        {
            get { return _path; }
            set
            {
                if (SQLiteDBUtills.SetValue(this, "Path", value))
                {
                    _path = value;
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
        private DbSubDepartment _department { get; set; }
        public DbSubDepartment Department
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
        private DbProject _project { get; set; }
        public DbProject Project
        {
            get { return _project; }
            set
            {
                if (SQLiteDBUtills.SetValue(this, "Project", value.Id))
                {
                    _project = value;
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
