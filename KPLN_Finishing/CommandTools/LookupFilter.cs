using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Finishing.CommandTools
{
    public class LookupFilter : AbstractFinishingFilter
    {
        private string Parameter;
        public LookupFilter(string parameter)
        { Parameter = parameter; }
        public override string GetValue(Room room)
        { return room.LookupParameter(Parameter).AsString(); }
    }
}
