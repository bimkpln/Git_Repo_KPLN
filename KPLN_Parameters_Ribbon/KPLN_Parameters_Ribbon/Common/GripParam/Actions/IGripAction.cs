using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Parameters_Ribbon.Common.GripParam.Actions
{
    internal interface IGripAction
    {
        string Name { get; }
        bool Compilte();
    }
}
