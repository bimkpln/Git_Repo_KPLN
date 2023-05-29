using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_ModelChecker_User.Common
{
    internal class MEPElementData
    {
        public MEPElementData(Element elem)
        {
            CurrentElement = elem;
        }

        public Element CurrentElement { get; }
        
        public Solid CurrentSolid { get; set; }

        public BoundingBoxXYZ CurrentBBox { get; private set; }
    }
}
