using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using KPLN_RevitMcpBridge.Server;

namespace Autodesk.Revit.DB
{
    public enum BuiltInCategory { OST_RoomTags = -1, OST_Rooms = -2 }
    public enum BuiltInParameter { VIEW_PHASE, ROOM_PHASE_ID }
    public enum ViewType { FloorPlan, CeilingPlan, AreaPlan }
    public enum SpatialElementTagOrientation { Horizontal, Vertical }
    public class ElementId
    {
        public long Value;
        public ElementId(long value) { Value = value; }
        public ElementId(BuiltInCategory value) { Value = (long)value; }
        public static bool operator ==(ElementId a, ElementId b) => ReferenceEquals(a, b) || (!ReferenceEquals(a, null) && !ReferenceEquals(b, null) && a.Value == b.Value);
        public static bool operator !=(ElementId a, ElementId b) => !(a == b);
        public override bool Equals(object other) => this == other as ElementId;
        public override int GetHashCode() => Value.GetHashCode();
    }
    public class XYZ
    {
        public double X, Y, Z;
        public XYZ(double x, double y, double z) { X = x; Y = y; Z = z; }
        public static XYZ BasisZ => new XYZ(0, 0, 1);
        public bool IsAlmostEqualTo(XYZ p) => Math.Abs(X-p.X)+Math.Abs(Y-p.Y)+Math.Abs(Z-p.Z) < 1e-9;
    }
    public class UV { public double U, V; public UV(double u, double v) { U=u; V=v; } }
    public class Transform
    {
        public XYZ BasisZ = XYZ.BasisZ;
        public double A = 1, B, C, D = 1, X, Y, Z;
        public XYZ OfPoint(XYZ p) => new XYZ(A*p.X+B*p.Y+X, C*p.X+D*p.Y+Y, p.Z+Z);
        public Transform Inverse
        {
            get
            {
                double det = A*D-B*C;
                return new Transform { A=D/det, B=-B/det, C=-C/det, D=A/det,
                    X=(-D*X+B*Y)/det, Y=(C*X-A*Y)/det, Z=-Z };
            }
        }
    }
    public class BoundingBoxXYZ { public Transform Transform = new Transform(); public XYZ Min, Max; }
    public class Category { public ElementId Id; }
    public class Parameter { public ElementId Value; public ElementId AsElementId() => Value; }
    public class LocationPoint { public XYZ Point; }
    public class Element
    {
        public ElementId Id, OwnerViewId, LevelId, TypeId;
        public string UniqueId, Name;
        public Category Category;
        public bool Hidden;
        public Dictionary<BuiltInParameter, Parameter> Parameters = new Dictionary<BuiltInParameter, Parameter>();
        public ElementId GetTypeId() => TypeId;
        public bool IsHidden(ViewPlan view) => Hidden;
        public Parameter get_Parameter(BuiltInParameter p) => Parameters.ContainsKey(p) ? Parameters[p] : null;
    }
    public class Level : Element { public double ProjectElevation; }
    public class Phase : Element { }
    public class ViewPlan : Element
    {
        public bool IsTemplate, CropBoxActive, TagsHidden;
        public ViewType ViewType;
        public Level GenLevel;
        public BoundingBoxXYZ CropBox;
        public CropManager Crop = new CropManager();
        public bool GetCategoryHidden(ElementId id) => TagsHidden;
        public CropManager GetCropRegionShapeManager() => Crop;
    }
    public class CropManager : IDisposable
    {
        public bool ShapeSet;
        public int NumberOfSplitRegions = 1;
        public void Dispose() { }
    }
    public class RevitLinkInstance : Element
    {
        public Document LinkedDocument;
        public Transform Transform = new Transform();
        public Document GetLinkDocument() => LinkedDocument;
        public Transform GetTransform() => Transform;
    }
    public class RevitLinkType : Element
    {
        public Dictionary<ElementId, ElementId> Map = new Dictionary<ElementId, ElementId>();
        public IDictionary<ElementId, ElementId> GetPhaseMap() => Map;
    }
    public class LinkElementId
    {
        public ElementId LinkInstanceId, LinkedElementId;
        public LinkElementId(ElementId link, ElementId room) { LinkInstanceId=link; LinkedElementId=room; }
    }
    public class Document
    {
        public List<Element> Elements = new List<Element>();
        public bool ReturnNull, Orphan;
        public Creator Create => new Creator(this);
        public Element GetElement(string id) => Elements.FirstOrDefault(e => e.UniqueId == id);
        public Element GetElement(ElementId id) => Elements.FirstOrDefault(e => e.Id == id);
        public void Regenerate() { }
        public T Add<T>(T e, long id) where T : Element { e.Id=new ElementId(id); e.UniqueId="uid-"+id; e.Name=e.UniqueId; Elements.Add(e); return e; }
    }
    public class Creator
    {
        private readonly Document doc;
        public Creator(Document document) { doc=document; }
        public Architecture.RoomTag NewRoomTag(LinkElementId target, UV point, ElementId view)
        {
            if (doc.ReturnNull) return null;
            return doc.Add(new Architecture.RoomTag { TaggedRoomId=target, OwnerViewId=view,
                TagHeadPosition=new XYZ(point.U, point.V, 0), IsOrphaned=doc.Orphan,
                Category=new Category { Id=new ElementId(BuiltInCategory.OST_RoomTags) } }, 10000+doc.Elements.Count);
        }
    }
    public class FilteredElementCollector : IEnumerable<Element>
    {
        private IEnumerable<Element> elements;
        public FilteredElementCollector(Document doc) { elements=doc.Elements; }
        public FilteredElementCollector OfCategory(BuiltInCategory cat) { elements=elements.Where(e => e.Category?.Id == new ElementId(cat)); return this; }
        public FilteredElementCollector OfClass(Type t) { elements=elements.Where(t.IsInstanceOfType); return this; }
        public FilteredElementCollector WhereElementIsNotElementType() => this;
        public IEnumerator<Element> GetEnumerator() => elements.GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
namespace Autodesk.Revit.DB.Architecture
{
    public class Room : Element { public object Location; public double Area = 10; public string Number = "101"; }
    public class RoomTagType : Element { }
    public class RoomTag : Element
    {
        public bool IsOrphaned, HasLeader;
        public LinkElementId TaggedRoomId;
        public RoomTagType RoomTagType;
        public SpatialElementTagOrientation TagOrientation;
        public XYZ TagHeadPosition;
    }
}
namespace KPLN_RevitMcpBridge.Services
{
    internal static class RevitService
    {
        internal static long IdValue(Autodesk.Revit.DB.ElementId id) => id.Value;
        internal static object ElementInfo(Autodesk.Revit.DB.Element e) => new { unique_id=e.UniqueId };
    }
    internal static class DocumentationReadService
    {
        internal static T Require<T>(Autodesk.Revit.DB.Document doc, IDictionary<string, object> row, string key) where T : Autodesk.Revit.DB.Element
            => doc.GetElement(row.Text(key, true)) as T ?? throw new BridgeException("wrong_element_type", key);
        internal static double[] Point(Autodesk.Revit.DB.XYZ p) => new[] {p.X*304.8, p.Y*304.8, p.Z*304.8};
    }
}
