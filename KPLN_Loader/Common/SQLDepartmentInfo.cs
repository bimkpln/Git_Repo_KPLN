using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Loader.Common
{
    public class SQLDepartmentInfo
    {
        public int Id { get; }
        public string Name { get; }
        public string Code { get; }
        public SQLDepartmentInfo(int id, string name, string code)
        {
            Id = id;
            Name = name;
            Code = code;
        }
    }
}
