using System;
using System.Collections.Generic;
namespace Autodesk.Revit.DB
{
    enum BuiltInCategory { OST_GenericAnnotation = -2000150 }
    class ElementId { public long Value; public ElementId(long value) { Value = value; } }
    class Category { public ElementId Id; }
    class XYZ { public double X, Y, Z; }
    class Location { }
    class LocationPoint : Location
    {
        public XYZ Point = new XYZ();
        public bool Unsupported;
        public double Rotation => Unsupported ? throw new ApplicationException("Unable to extract rotation from family instance.") : 0.5;
    }
    class LocationCurve : Location { public Curve Curve; }
    class Curve { public bool IsBound = true; public double Length = 1; public XYZ GetEndPoint(int index) => new XYZ { X = index }; }
    class Transform
    {
        public XYZ Origin = new XYZ();
        public XYZ BasisX = new XYZ { X = 1 }, BasisY = new XYZ { Y = 1 }, BasisZ = new XYZ { Z = 1 };
    }
    class BoundingBoxXYZ { public XYZ Min = new XYZ(), Max = new XYZ(); public Transform Transform = new Transform(); }
    class Element
    {
        public ElementId Id = new ElementId(1), OwnerViewId = new ElementId(2), TypeId = new ElementId(3);
        public Category Category;
        public object Parameters;
        public Location Location;
        public BoundingBoxXYZ Box;
        public bool ThrowBounds;
        public ElementId GetTypeId() => TypeId;
        public BoundingBoxXYZ get_BoundingBox(View view)
        {
            if (ThrowBounds) throw new InvalidOperationException("bounds unsupported");
            return Box;
        }
    }
    class View : Element { }
    class FamilyInstance : Element { }
    class FamilySymbol : Element { public string FamilyName; }
    class TextNote : Element { public string Text; public XYZ Coord = new XYZ(); }
    class Document
    {
        public Dictionary<long, Element> Elements = new Dictionary<long, Element>();
        public Element GetElement(ElementId id) { Element result; Elements.TryGetValue(id.Value, out result); return result; }
    }
}
namespace KPLN_RevitMcpBridge.Services
{
    using Autodesk.Revit.DB;
    static class RevitService
    {
        public static long IdValue(ElementId id) => id.Value;
        public static object ElementInfo(Element e) => new { id = e.Id.Value.ToString() };
    }
    static class ParameterService { public static object Read(Element e) => e.Parameters; }
}
