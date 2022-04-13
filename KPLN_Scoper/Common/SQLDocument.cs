using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Scoper.Common
{
    public class SQLDocument
    {
        public int Id { get; set; }
        public string Path { get; set; }
        public string Name { get; set; }
        public int Department { get; set; }
        public int Project { get; set; }
        public string Code { get; set; }
        public SQLDocument(int id, string path, string name, int department, int project, string code)
        {
            Id = id;
            Path = path;
            Name = name;
            Department = department;
            Project = project;
            Code = code;
        }
    }
}
