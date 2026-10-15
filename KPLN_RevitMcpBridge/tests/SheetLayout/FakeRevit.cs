using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace Autodesk.Revit.DB
{
    public class ElementId
    {
        public readonly int Value;
        public ElementId(int value) { Value = value; }
        public static readonly ElementId InvalidElementId = new ElementId(-1);
        public static bool operator ==(ElementId a, ElementId b) => ReferenceEquals(a, b) || (!ReferenceEquals(a, null) && !ReferenceEquals(b, null) && a.Value == b.Value);
        public static bool operator !=(ElementId a, ElementId b) => !(a == b);
        public override bool Equals(object o) => o is ElementId && this == (ElementId)o;
        public override int GetHashCode() => Value;
    }
    public class XYZ
    {
        public double X, Y, Z;
        public XYZ(double x, double y, double z) { X = x; Y = y; Z = z; }
        public double DistanceTo(XYZ b) => Math.Sqrt((X-b.X)*(X-b.X)+(Y-b.Y)*(Y-b.Y)+(Z-b.Z)*(Z-b.Z));
        public static XYZ operator +(XYZ a, XYZ b) => new XYZ(a.X+b.X,a.Y+b.Y,a.Z+b.Z);
        public static XYZ operator -(XYZ a, XYZ b) => new XYZ(a.X-b.X,a.Y-b.Y,a.Z-b.Z);
        public static XYZ operator *(XYZ a, double k) => new XYZ(a.X*k,a.Y*k,a.Z*k);
        public double DotProduct(XYZ b) => X*b.X+Y*b.Y+Z*b.Z;
    }
    public class Transform
    {
        public XYZ Origin = new XYZ(0,0,0);
        public bool Rotate;
        public XYZ OfPoint(XYZ p) => Origin + (Rotate ? new XYZ(-p.Y,p.X,p.Z) : p);
    }
    public class BoundingBoxXYZ { public XYZ Min, Max; public Transform Transform = new Transform(); }
    public class Outline
    {
        public XYZ MinimumPoint, MaximumPoint;
        public Outline(XYZ min, XYZ max) { MinimumPoint=min; MaximumPoint=max; }
    }
    public class UV { public double U,V; public UV(double u,double v) { U=u; V=v; } }
    public class BoundingBoxUV { public UV Min,Max; }
    public enum BuiltInParameter { VIEWPORT_ATTR_SHOW_LABEL }
    public class Parameter { public int Value; public int AsInteger() => Value; }
    public class Element
    {
        public ElementId Id, OwnerViewId, GroupId = ElementId.InvalidElementId, TypeId;
        public string UniqueId;
        public bool Pinned;
        public BoundingBoxXYZ Box;
        public Parameter Label;
        public virtual BoundingBoxXYZ get_BoundingBox(ViewSheet sheet) => Box;
        public ElementId GetTypeId() => TypeId;
        public Parameter get_Parameter(BuiltInParameter p) => Label;
    }
    public class ViewSheet : Element
    {
        public bool IsPlaceholder;
        public BoundingBoxUV Outline;
        public List<ElementId> Viewports = new List<ElementId>();
        public ICollection<ElementId> GetAllViewports() => Viewports;
    }
    public class Viewport : Element
    {
        public int Rotation;
        public bool IgnoreLabelMove;
        private XYZ labelOffset = new XYZ(0,0,0);
        public XYZ LabelOffset
        {
            get => labelOffset;
            set
            {
                if (IgnoreLabelMove) return;
                var delta = value-labelOffset;
                if (LabelBounds!=null) LabelBounds=new Outline(LabelBounds.MinimumPoint+delta,LabelBounds.MaximumPoint+delta);
                labelOffset=value;
            }
        }
        public double LabelLineLength { get; set; }
        public Outline Bounds, LabelBounds;
        public bool IgnoreMove;
        public XYZ GetBoxCenter() => new XYZ((Bounds.MinimumPoint.X+Bounds.MaximumPoint.X)/2, (Bounds.MinimumPoint.Y+Bounds.MaximumPoint.Y)/2,17);
        public Outline GetBoxOutline() => Bounds;
        public Outline GetLabelOutline() => LabelBounds;
        public void SetBoxCenter(XYZ p)
        {
            if (IgnoreMove) return;
            var c=GetBoxCenter(); var delta = new XYZ(p.X-c.X,p.Y-c.Y,0);
            Bounds = new Outline(Bounds.MinimumPoint+delta,Bounds.MaximumPoint+delta);
            if (LabelBounds!=null) LabelBounds=new Outline(LabelBounds.MinimumPoint+delta,LabelBounds.MaximumPoint+delta);
        }
    }
    public class ScheduleSheetInstance : Element
    {
        public bool IsTitleblockRevisionSchedule;
        private XYZ point = new XYZ(0,0,0);
        public XYZ Point { get => point; set { if(Box!=null) { Box.Min+=value-point; Box.Max+=value-point; } point=value; } }
    }
    public class Document
    {
        public List<Element> Elements = new List<Element>();
        public Element GetElement(ElementId id) => Elements.FirstOrDefault(e=>e.Id==id);
        public Element GetElement(string uid) => Elements.FirstOrDefault(e=>e.UniqueId==uid);
        public void Regenerate() { }
    }
    public class FilteredElementCollector : IEnumerable<Element>
    {
        private IEnumerable<Element> elements;
        public FilteredElementCollector(Document doc) { elements=doc.Elements; }
        public FilteredElementCollector OfClass(Type type) { elements=elements.Where(type.IsInstanceOfType); return this; }
        public IEnumerator<Element> GetEnumerator() => elements.GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
namespace KPLN_RevitMcpBridge.Services
{
    using Autodesk.Revit.DB;
    using KPLN_RevitMcpBridge.Server;
    internal static class DocumentationReadService
    {
        internal static T Require<T>(Document doc, IDictionary<string,object> input,string key) where T:Element
        {
            var e=doc.GetElement(input.Text(key,true)) as T;
            if(e==null) throw new BridgeException("wrong_element_type",key);
            return e;
        }
    }
    internal static class DocumentationWriteService
    {
        internal static XYZ Point(IDictionary<string,object> input,string key)
        {
            var a=input.Get(key) as IList;
            if(a==null||a.Count!=3) throw new BridgeException("invalid_point",key);
            return new XYZ(Json.Double(a[0])/304.8,Json.Double(a[1])/304.8,Json.Double(a[2])/304.8);
        }
    }
}
