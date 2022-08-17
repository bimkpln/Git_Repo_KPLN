using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Finishing.CommandTools
{
    public abstract class AbstractOnIdlingCommand
    {
        public virtual void Execute(UIApplication uiapp)
        { throw new ArgumentException(); }
    }
}
