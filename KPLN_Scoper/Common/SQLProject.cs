using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Scoper.Common
{
    public class SQLProject
    {
        public int Id { get; }
        public string Name { get; }
        public string Code { get; }
        public string Keys { get; }
        public SQLProject(int id, string name, string code, string keys)
        {
            Id = id;
            Name = name;
            Code = code;
            Keys = keys;
        }
    }
}
