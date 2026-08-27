using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using KPLN_MEPBender.Common;

namespace KPLN_MEPBender.Services.Routing
{
    internal sealed class MepCurveFactory
    {
        private static readonly double MinSegmentLength = UnitConvert.MmToInternal(10);

        public MEPCurve Create(Document doc, MEPCurve source, XYZ start, XYZ end)
        {
            if (start.DistanceTo(end) < MinSegmentLength)
                return null;

            ElementId typeId = source.GetTypeId();
            ElementId levelId = source.ReferenceLevel?.Id ?? source.LevelId;

            Pipe pipe = source as Pipe;
            if (pipe != null)
            {
                return Pipe.Create(
                    doc,
                    GetSystemTypeId(source, BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM),
                    typeId,
                    levelId,
                    start,
                    end);
            }

            Duct duct = source as Duct;
            if (duct != null)
            {
                return Duct.Create(
                    doc,
                    GetSystemTypeId(source, BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM),
                    typeId,
                    levelId,
                    start,
                    end);
            }

            CableTray cableTray = source as CableTray;
            if (cableTray != null)
                return CableTray.Create(doc, typeId, start, end, levelId);

            return null;
        }

        private ElementId GetSystemTypeId(MEPCurve source, BuiltInParameter systemTypeParameter)
        {
            ElementId systemTypeId = source.MEPSystem?.GetTypeId();
            if (systemTypeId != null && systemTypeId != ElementId.InvalidElementId)
                return systemTypeId;

            Parameter parameter = source.get_Parameter(systemTypeParameter);
            if (parameter != null && parameter.StorageType == StorageType.ElementId)
                return parameter.AsElementId();

            return ElementId.InvalidElementId;
        }
    }
}
