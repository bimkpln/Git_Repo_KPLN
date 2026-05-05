using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Library_DBWorker.Core
{
    [Serializable]
    public class Worker_Error: Exception
    {
        public Worker_Error(string msg) : base($"\n\tОшибка модуля [KPLN_Library_DBWorker]:\n{msg}")
        {
        }
    }
}
