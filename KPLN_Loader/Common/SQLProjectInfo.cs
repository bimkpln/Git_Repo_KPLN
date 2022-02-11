using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Loader.Common
{
    public class SQLProjectInfo
    {
        public int Id { get;}
        public string Name { get; }
        public SQLProjectInfo(int id, string name)
        {
            Id = id;
            Name = name;
        }
    }
}
