using KPLN_DataBase.Collections;
using System.Collections.ObjectModel;

namespace KPLN_DataBase.Controll
{
    public class DbUserInfo
    {
        public DbUserInfo(string systemName, string name, string family, string surname, int department, string status, string connection, string sex, ObservableCollection<DbDepartment> departments)
        {
            SystemName = systemName;
            Name = name;
            Family = family;
            Surname = surname;
            foreach (DbDepartment d in departments)
            {
                if (d.Id == department)
                { Department = d; }
            }
            Status = status;
            Connection = connection;
            Sex = sex;
        }
        public string SystemName { get; }
        public string Name { get; }
        public string Family { get; }
        public string Surname { get; }
        public DbDepartment Department { get; }
        public string Status { get; }
        public string Connection { get; }
        public string Sex { get; }
    }
}
