using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Parameters_Ribbon.Common.Tools
{
    internal class SolidComparerTool : IEqualityComparer<Solid>
    {
        public bool Equals(Solid x, Solid y)
        {
            return x == y;
        }

        public int GetHashCode(Solid obj)
        {
            return obj.GetHashCode();
        }
    }
}
