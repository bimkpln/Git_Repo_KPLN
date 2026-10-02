using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_CalculateTEP.Common;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using PC = KPLN_CalculateTEP.Common.Geometry.TepClipper.Clipper;
using CP = KPLN_CalculateTEP.Common.Geometry.TepClipper.IntPoint;
using CT = KPLN_CalculateTEP.Common.Geometry.TepClipper.ClipType;
using PT = KPLN_CalculateTEP.Common.Geometry.TepClipper.PolyType;
using PF = KPLN_CalculateTEP.Common.Geometry.TepClipper.PolyFillType;

using TepClipper = KPLN_CalculateTEP.Common.Geometry.TepClipper;
using KPLN_CalculateTEP.Common.Methodologies;
namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public class Source
        {
            public string Key {get;set;} public string Name {get;set;} public string Mode {get;set;}="include";
            public string Profile {get;set;}="residential"; public bool Loaded {get {return Document!=null;}}
            public string LoadError {get;set;}
            internal Document Document; internal Transform Transform; internal ElementId RootLink;
            internal Document ParentDocument; internal ElementId LinkTypeId;
            internal List<Element> Elements=new List<Element>();
        }
        internal class Record
        {
            internal Source Source; internal Element Element; internal LevelSetting Level;
            internal string Building,Section,Profile,BuildingClass,Role,Part,Apartment,Vertical;
            internal double? Factor; internal bool Manual; internal bool? Override; internal bool SingleStorey; internal bool BuildingIncluded=true;
            internal string Key {get {return Source.Key+"/"+Element.UniqueId;}}
            internal double Z {get {return Level==null?0:Level.Elevation;}}
            internal Record Copy() {return (Record)MemberwiseClone();}
        }
    }
}
