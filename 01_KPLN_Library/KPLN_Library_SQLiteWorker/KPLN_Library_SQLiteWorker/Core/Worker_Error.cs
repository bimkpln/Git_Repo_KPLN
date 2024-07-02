using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Library_SQLiteWorker.Core
{
    [Serializable]
    public class Worker_Error: Exception
    {
        public Worker_Error(string msg) : base($"\n\tОшибка модуля [SQLiteWorker]:\n{msg}")
        {
        }
    }
}
