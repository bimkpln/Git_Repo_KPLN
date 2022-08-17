using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Finishing.CommandTools
{
    public abstract class AbstractFinishingFilter
    {
        public virtual string GetValue(Room room)
        {
            return null;
        }
    }
}
