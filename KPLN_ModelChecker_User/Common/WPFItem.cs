using Autodesk.Revit.DB;
using KPLN_ModelChecker_Lib.Common;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using static KPLN_ModelChecker_User.Common.Collections;

namespace KPLN_ModelChecker_User.Common
{
    public class WPFItem
    {
        public SolidColorBrush Fill { get; set; }
        
        public BoundingBoxXYZ Box { get; set; }
        
        public XYZ Centroid { get; set; }
        
        private IList<ElementEntity> ElementEntities { get; set; }

        public WPFItem(IList<ElementEntity> elementEntities)
        {
            elementEntities = ElementEntities;
        }
    }

}
