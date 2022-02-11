using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Loader.Common
{
    public class SQLModuleInfo
    {
        public int Id { get; }
        public int Department { get; }
        public string Version { get; }
        public string Path { get; }
        public string Name { get; }
        public SQLModuleInfo(int id, int dep, string ver, string path, string name)
        {
            Id = id;
            Department = dep;
            Version = ver;
            Path = path;
            Name = name;

        }
    }
}
