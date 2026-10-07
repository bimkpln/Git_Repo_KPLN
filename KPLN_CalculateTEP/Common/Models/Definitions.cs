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
        public enum Indicator { Gns, GnsResidential, GnsLivingPart, GnsNonlivingPart, GnsNonresidential, Footprint,
            Volume, VolumeAbove, VolumeBelow, Gross, GrossAbove, GrossBelow, NnpEmbedded, NnpSeparate,
            Np, NpResidential, NpNonresidential, PublicCalculated, ApartmentsTotal, ApartmentsHeated,
            PublicRooms, ApartmentsCount, ParkingCount, Storeys, Floors }
        public class Choice
        {
            public string Key { get; set; }
            public string Label { get; set; }
            public string Description { get; set; }
            public Choice() { }
            public Choice(string key, string label, string description) { Key = key; Label = label; Description = description; }
        }
        public class Metric : System.ComponentModel.INotifyPropertyChanged
        {
            public string Key { get; set; }
            public string Name { get; set; }
            public string Unit { get; set; }
            public string Description { get; set; }
            public bool Enabled { get; set; }
            public string Geometry { get; set; }
            public string AreaScheme { get; set; }
            private string contourMode="current";
            public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
            public string ContourMode {get{return contourMode;}set{if(contourMode==value)return;contourMode=value;PropertyChanged?.Invoke(this,new System.ComponentModel.PropertyChangedEventArgs(nameof(ContourMode)));}}
            public string WallSelectionMode {get;set;} = "exterior";
            public string WallParameter {get;set;}
            public string WallValue {get;set;}
            public string WallBoundary {get;set;} = "auto";
            public string WallCutHeight {get;set;} = "0";
            public List<string> SelectedWalls {get;set;} = new List<string>();
            public string ContourExclusions {get;set;} = "normative";
            public string VolumeMaskHeight {get;set;} = "actual";
        }
        public class ContourSketch
        {
            public string Metric {get;set;} public string ElementUniqueId {get;set;}
            public string LevelKey {get;set;} public string LevelName {get;set;}
            public string Kind {get;set;} = "base";
            public string Building {get;set;} public string Section {get;set;} = "";
            public string Profile {get;set;} = "residential";
            public string BuildingClass {get;set;} = "residential";
            public string Height {get;set;} public string BottomOffset {get;set;} = "0";
            public string ElementLabel {get;set;}
        }
        public class ParameterMap : System.ComponentModel.INotifyPropertyChanged
        {
            public string Key {get;set;} public string Title {get;set;} public string Name {get;set;} public string Description {get;set;}
            private bool isRequired;
            public bool IsRequired {get{return isRequired;}set{if(isRequired==value)return;isRequired=value;PropertyChanged?.Invoke(this,new System.ComponentModel.PropertyChangedEventArgs(nameof(IsRequired)));}}
            public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
        }
        public class Rule
        {
            public bool Enabled {get;set;} = true;
            public string Metric {get;set;} = "all";
            public string Category {get;set;}
            public string Parameter {get;set;} = "@Name";
            public string Operation {get;set;} = "contains";
            public string Value {get;set;}
            public string Role {get;set;} = "unknown";
            public string Part {get;set;} = "auto";
            public string Action {get;set;} = "classify";
            public double? Coefficient {get;set;}
        }
        public class SourceOption { public string Key {get;set;} public string Mode {get;set;} public string Profile {get;set;} }
        public class BuildingMap
        {
            public string Source {get;set;} = "*";
            public string MatchValue {get;set;} = "*";
            public string Building {get;set;}
            public string Profile {get;set;} = "residential";
            public string Class {get;set;} = "residential";
            public bool Include {get;set;} = true;
        }
        public class LevelSetting : System.ComponentModel.INotifyPropertyChanged
        {
            public string Key {get;set;} public string Source {get;set;} public string Name {get;set;}
            public string Building {get;set;} public string Section {get;set;}
            public double Elevation {get;set;} public bool Include {get;set;} = true;
            private string kind="normal",above="auto";
            public string Kind {get{return kind;}set{if(kind==value)return;kind=value;NotifySlabRequirement("Kind");}}
            public string Above {get{return above;}set{if(above==value)return;above=value;NotifySlabRequirement("Above");}}
            [System.Runtime.Serialization.IgnoreDataMember]
            public bool TopSlabRequired {get{return Kind=="basement"&&Above=="auto";}}
            private void NotifySlabRequirement(string property)
            {foreach(var name in new[]{property,"TopSlabRequired","TopSlabSummary","HeightRequired","RoofAreaRequired","RoofRatioRequired"})NotifyValue(name);}
            public string TopSlab {get;set;} public string Height {get;set;} public string RoofRatio {get;set;} public string RoofArea {get;set;}
            public string TopSlabMode {get;set;}
            public string TopSlabChoice {get;set;}
            public string HeightMode {get;set;}
            public string RoofAreaMode {get;set;}
            public string RoofRatioMode {get;set;}
            public List<string> RoofSources {get;set;}=new List<string>();
            private string normativeProfile="residential";
            private bool hasRoofObjects;
            private bool hasHighObjects,hasPublicObjects;
            [System.Runtime.Serialization.IgnoreDataMember]
            public bool HasHighObjects {get{return hasHighObjects;}set{hasHighObjects=value;NotifySlabRequirement("HasHighObjects");}}
            [System.Runtime.Serialization.IgnoreDataMember]
            public bool HasPublicObjects {get{return hasPublicObjects;}set{hasPublicObjects=value;NotifySlabRequirement("HasPublicObjects");}}
            [System.Runtime.Serialization.IgnoreDataMember]
            public string NormativeProfile {get{return normativeProfile;}set{normativeProfile=value;NotifySlabRequirement("NormativeProfile");}}
            [System.Runtime.Serialization.IgnoreDataMember]
            public bool HasRoofObjects {get{return hasRoofObjects;}set{hasRoofObjects=value;NotifySlabRequirement("HasRoofObjects");}}
            [System.Runtime.Serialization.IgnoreDataMember]
            public bool HeightRequired {get{return Kind=="void"&&(NormativeProfile=="public"||HasPublicObjects)||(Kind=="roof"||HasRoofObjects)&&((NormativeProfile??"").StartsWith("high-")||HasHighObjects);}}
            [System.Runtime.Serialization.IgnoreDataMember]
            public bool RoofAreaRequired {get{return (Kind=="roof"||HasRoofObjects)&&((NormativeProfile??"").StartsWith("high-")||HasHighObjects);}}
            [System.Runtime.Serialization.IgnoreDataMember]
            public bool RoofRatioRequired {get{return (Kind=="roof"||HasRoofObjects)&&(NormativeProfile=="public"||NormativeProfile=="high-public"||HasPublicObjects);}}
            private string heightSummary,roofAreaSummary,roofRatioSummary;
            [System.Runtime.Serialization.IgnoreDataMember]
            public string HeightSummary {get{return heightSummary??"Задать";}set{heightSummary=value;NotifyValue("HeightSummary");}}
            [System.Runtime.Serialization.IgnoreDataMember]
            public string RoofAreaSummary {get{return roofAreaSummary??"Задать";}set{roofAreaSummary=value;NotifyValue("RoofAreaSummary");}}
            [System.Runtime.Serialization.IgnoreDataMember]
            public string RoofRatioSummary {get{return roofRatioSummary??"Задать";}set{roofRatioSummary=value;NotifyValue("RoofRatioSummary");}}
            private void NotifyValue(string name){PropertyChanged?.Invoke(this,new System.ComponentModel.PropertyChangedEventArgs(name));}
            private string topSlabSummary;
            [System.Runtime.Serialization.IgnoreDataMember]
            public string TopSlabSummary {get{return !TopSlabRequired?"":string.IsNullOrEmpty(topSlabSummary)?"Задать":topSlabSummary;}
                set{topSlabSummary=value;PropertyChanged?.Invoke(this,new System.ComponentModel.PropertyChangedEventArgs("TopSlabSummary"));}}
            public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
            public double AbsoluteElevationMeters {get;set;}
            public string DatumDisplayName {get{return Source+" / "+Name+" (абс. "+AbsoluteElevationMeters.ToString("0.000")+" м)";}}
            public double ElevationMeters {get{return Elevation*.3048;}}
            public string DisplayName {get{return Source+" / "+Name+" ("+ElevationMeters.ToString("0.000")+" м)";}}
        }
        public class Correction
        {
            public string Metric {get;set;} public string Action {get;set;}
            public string Source {get;set;} public string Element {get;set;}
            public string Value {get;set;} public string Reason {get;set;}
            public string Author {get;set;} public string Date {get;set;} public string Previous {get;set;}
        }
    }
}
