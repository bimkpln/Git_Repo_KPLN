using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Loader.Common
{
    public class SQLUserInfo
    {
        public string SystemName { get; }
        public string Name { get; }
        public string Surname { get; }
        public string Family { get; }
        public string Status { get; }
        public SQLDepartmentInfo Department { get; set; }
        public SQLUserInfo(string systemName, string name, string surname, string family, string status)
        {
            SystemName = systemName;
            Name = name;
            Surname = surname;
            Family = family;
            Status = status;
        }
    }
}
