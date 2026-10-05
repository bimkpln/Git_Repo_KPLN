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
        public class Issue
        {
            public string Code {get;set;} public string Severity {get;set;} public string Message {get;set;}
            public string Metric {get;set;} public string Source {get;set;} public string Building {get;set;}
            public string Element {get;set;} public string Action {get;set;}
            public string ParameterKey {get;set;} public string InputKind {get;set;}
            public string MetricName {get{return string.IsNullOrEmpty(Metric)?"Общая проверка":MetricLabel(Metric);}}
        }
        public class Detail
        {
            public string Metric {get;set;} public string SourceKey {get;set;} public string Source {get;set;}
            public string Building {get;set;} public string Section {get;set;} public string Level {get;set;}
            public double Elevation {get;set;} public string Element {get;set;} public string UniqueId {get;set;}
            public double? MeasurementElevation {get;set;}
            public string Purpose {get;set;} public string Apartment {get;set;} public string Profile {get;set;}
            public string Method {get;set;} public string Unit {get;set;} public double Raw {get;set;} public double Factor {get;set;}=1;
            public double Value {get;set;} public bool Excluded {get;set;} public bool Manual {get;set;}
            public string Reason {get;set;} public string ViewId {get;set;}
            public bool HasError {get;set;}
            public string MetricName {get{return MetricLabel(Metric);}}
            public string PurposeName {get{return RoleLabel(Purpose);}}
            internal Solid Shape; internal bool VolumeShape;
        }
        public class Summary
        {
            public string Key {get;set;} public string Name {get;set;} public double Value {get;set;}
            public string Unit {get;set;} public string Method {get;set;} public string Status {get;set;} public string Comment {get;set;}
            public bool NotCalculated {get;set;}
            public double? DisplayValue {get{return NotCalculated?(double?)null:Value;}}
            public string ValueText(int decimals) {return NotCalculated?"не рассчитано":Math.Round(Value,decimals,MidpointRounding.AwayFromZero).ToString("N"+decimals)+" "+Unit;}
        }
        public class Run
        {
            public string Id {get;set;}=Guid.NewGuid().ToString("N"); public string Date {get;set;}=DateTime.Now.ToString("s");
            public string GeometryVersion {get;set;} = TepCalculation.GeometryVersion;
            public bool? CreateViews {get;set;}
            public string Author {get;set;} public string Method {get;set;} public string Version {get;set;}=RulesVersion;
            public string Configuration {get;set;} public List<Summary> Summary {get;set;}=new List<Summary>();
            public List<Detail> Details {get;set;}=new List<Detail>(); public List<Issue> Issues {get;set;}=new List<Issue>();
            public List<Correction> Corrections {get;set;}=new List<Correction>();
            public void Issue(string code,string severity,string message,string metric="",string source="",string building="",string element="",string action="Проверьте настройки и исправьте исходные данные.")
            { Issues.Add(new Issue {Code=code,Severity=severity,Message=message,Metric=metric,Source=source,Building=building,Element=element,Action=action}); }
            public string TextReport(int decimals)
            {
                var b=new StringBuilder("Расчёт ТЭП | "+Date+" | "+Method+" | правила "+Version+(string.IsNullOrEmpty(GeometryVersion)?"":" | геометрия "+GeometryVersion)+(CreateViews.HasValue?" | проверочные виды: "+(CreateViews.Value?"включены":"выключены"):"")+"\n");
                foreach(var s in Summary)
                {
                    b.AppendLine();b.AppendLine(s.Name+": "+s.ValueText(decimals)+" | "+s.Status);
                    foreach(var g in Details.Where(d=>d.Metric==s.Key&&!d.Excluded).GroupBy(d=>new {d.Building,d.Section}).OrderBy(g=>g.Key.Building))
                    {
                        b.AppendLine("  "+g.Key.Building+(string.IsNullOrWhiteSpace(g.Key.Section)?"":" / секция "+g.Key.Section));
                        foreach(var floor in g.GroupBy(d=>Math.Round(d.Elevation,6)).OrderBy(x=>x.Key))
                            b.AppendLine("    "+string.Join(" / ",floor.Select(d=>d.Level).Distinct())+" ("+(floor.Key*.3048).ToString("+0.000;-0.000;0.000")+" м): "+Math.Round(floor.Sum(d=>d.Value),decimals,MidpointRounding.AwayFromZero).ToString("N"+decimals)+" "+s.Unit);
                    }
                    if(!string.IsNullOrEmpty(s.Comment)) b.AppendLine("  "+s.Comment);
                    foreach(var issue in Issues.Where(i=>Engine.IssueAffectsMetric(i,s.Key)&&
                        (i.Code=="SPATIAL_UNBOUNDED"||i.Code=="ROOM_FLOOR_SKIPPED")))
                        b.AppendLine("  "+issue.Source+": "+issue.Code+": "+issue.Message);
                }
                return b.ToString();
            }
        }
    }
}
