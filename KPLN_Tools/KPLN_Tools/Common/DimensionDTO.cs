using Autodesk.Revit.DB;

namespace KPLN_Tools.Common
{
    internal class DimensionDTO
    {
        public ElementId DimId { get; set; }
        
        public ElementId DimViewId { get; set; }

        public Curve DimCurve { get; set; }

        public XYZ DimLidEndPostion { get; set; }

        public ReferenceArray DimRefArray { get; set; }

        public DimensionType DimType { get; set; }
    }
}
