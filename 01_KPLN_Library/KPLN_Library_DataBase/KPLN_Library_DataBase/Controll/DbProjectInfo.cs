using KPLN_Library_DataBase.Collections;
using System.Collections.ObjectModel;

namespace KPLN_Library_DataBase.Controll
{
    public class DbProjectInfo
    {
        public DbProjectInfo(int id, string name, string users, string code, string keys, ObservableCollection<DbUser> collection)
        {
            Id = id;
            Name = name;
            Users = new ObservableCollection<DbUser>();
            Code = code;
            Keys = keys;
            foreach (string part in users.Split('*'))
            {
                foreach (DbUser user in collection)
                {
                    if (user.SystemName == part)
                    {
                        Users.Add(user);
                    }
                }
            }
        }
        public int Id { get; }
        public string Name { get; }
        public ObservableCollection<DbUser> Users { get; }
        public string Code { get; }
        public string Keys { get; }
    }
}
