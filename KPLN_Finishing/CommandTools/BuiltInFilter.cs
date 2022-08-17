using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Finishing.CommandTools
{
    public class BuiltInFilter : AbstractFinishingFilter
    {
        private BuiltInParameter Parameter;
        public BuiltInFilter(BuiltInParameter parameter)
        { Parameter = parameter; }
        public override string GetValue(Room room)
        { return room.get_Parameter(Parameter).AsString(); }
    }
}
