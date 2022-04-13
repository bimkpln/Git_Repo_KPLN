using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Scoper.Common
{
    class SQLDepartment
    {
        public int Id { get; }
        public string Name { get; }
        public string Code { get; }
        public string CodeUs { get; }
        public SQLDepartment(int id, string name, string code, string codeUs)
        {
            Id = id;
            Name = name;
            Code = code;
            CodeUs = codeUs;
        }
    }
}
