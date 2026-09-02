using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using KPLN_MEPBender.Common;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

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
                return CreateElectricalCurve(doc, cableTray, typeId, levelId, start, end);

            Conduit conduit = source as Conduit;
            if (conduit != null)
                return CreateElectricalCurve(doc, conduit, typeId, levelId, start, end);

            return null;
        }

        private MEPCurve CreateElectricalCurve(Document doc, MEPCurve source, ElementId typeId, ElementId levelId, XYZ start, XYZ end)
        {
            MEPCurve copiedCurve = CreateByCopy(doc, source, start, end);
            if (copiedCurve != null)
                return copiedCurve;

            CableTray cableTray = source as CableTray;
            if (cableTray != null)
                return CreateCableTray(doc, cableTray, typeId, levelId, start, end);

            Conduit conduit = source as Conduit;
            if (conduit != null)
                return CreateConduit(doc, conduit, typeId, levelId, start, end);

            return null;
        }

        private MEPCurve CreateByCopy(Document doc, MEPCurve source, XYZ start, XYZ end)
        {
            ICollection<ElementId> copiedElementIds;
            try
            {
                copiedElementIds = ElementTransformUtils.CopyElement(doc, source.Id, XYZ.Zero);
            }
            catch
            {
                return null;
            }

            ElementId copiedElementId = copiedElementIds.FirstOrDefault();
            if (copiedElementId == null || copiedElementId == ElementId.InvalidElementId)
                return null;

            MEPCurve copiedCurve = doc.GetElement(copiedElementId) as MEPCurve;
            LocationCurve copiedLocationCurve = copiedCurve?.Location as LocationCurve;
            if (copiedCurve == null || copiedLocationCurve == null)
                return null;

            try
            {
                copiedLocationCurve.Curve = Line.CreateBound(start, end);
                AlignCurveNormal(copiedCurve, source, start, end);
                return copiedCurve;
            }
            catch
            {
                doc.Delete(copiedElementId);
                return null;
            }
        }

        private CableTray CreateCableTray(Document doc, CableTray source, ElementId typeId, ElementId levelId, XYZ start, XYZ end)
        {
            CableTray cableTray = CableTray.Create(doc, typeId, start, end, levelId);
            AlignCurveNormal(cableTray, source, start, end);
            return cableTray;
        }

        private Conduit CreateConduit(Document doc, Conduit source, ElementId typeId, ElementId levelId, XYZ start, XYZ end)
        {
            Conduit conduit = Conduit.Create(doc, typeId, start, end, levelId);
            AlignCurveNormal(conduit, source, start, end);
            return conduit;
        }

        private void AlignCurveNormal(MEPCurve target, MEPCurve source, XYZ start, XYZ end)
        {
            if (target == null || source == null)
                return;

            XYZ curveDirection = (end - start).Normalize();
            XYZ normal = GetPerpendicularNormal(GetCurveNormal(source), curveDirection);
            if (normal == null)
                normal = GetPerpendicularNormal(XYZ.BasisZ, curveDirection);
            if (normal == null)
                normal = GetPerpendicularNormal(XYZ.BasisX, curveDirection);
            if (normal == null)
                normal = GetPerpendicularNormal(XYZ.BasisY, curveDirection);

            if (normal == null)
                return;

            SetCurveNormal(target, normal);
        }

        private XYZ GetCurveNormal(MEPCurve curve)
        {
            PropertyInfo property = curve?.GetType().GetProperty("CurveNormal");
            if (property == null || property.PropertyType != typeof(XYZ))
                return null;

            try
            {
                return property.GetValue(curve, null) as XYZ;
            }
            catch
            {
                return null;
            }
        }

        private void SetCurveNormal(MEPCurve curve, XYZ normal)
        {
            PropertyInfo property = curve?.GetType().GetProperty("CurveNormal");
            if (property == null || !property.CanWrite || property.PropertyType != typeof(XYZ))
                return;

            try
            {
                property.SetValue(curve, normal, null);
            }
            catch
            {
                // Some families/segment orientations are locked by Revit.
            }
        }

        private XYZ GetPerpendicularNormal(XYZ candidateNormal, XYZ curveDirection)
        {
            if (candidateNormal == null || curveDirection == null)
                return null;

            XYZ normal = candidateNormal - curveDirection * candidateNormal.DotProduct(curveDirection);
            return normal.GetLength() < 0.001 ? null : normal.Normalize();
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
