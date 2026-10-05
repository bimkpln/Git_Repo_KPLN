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
        public const string GeometryVersion = "2026.10.05 / conics-room-masks-5";
        public const string RulesVersion = "2026.08.26 / 2.0";
        private static readonly Dictionary<string,string> MetricLabels=Catalog().ToDictionary(m=>m.Key,m=>m.Name);
        private static readonly Dictionary<string,string> RoleLabels=Roles().ToDictionary(r=>r.Key,r=>r.Label);
        private static string MetricLabel(string key){string label;return key!=null&&MetricLabels.TryGetValue(key,out label)?label:key??"";}
        private static string RoleLabel(string key){string label;return key!=null&&RoleLabels.TryGetValue(key,out label)?label:key??"";}
    }
}
