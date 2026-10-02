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
        public partial class Engine
        {
            private readonly UIApplication app;
            private readonly UIDocument ui;
            private readonly Document doc;
            private readonly HashSet<long> selection;
            private static readonly Guid StorageId=new Guid("DAD13984-7185-4C89-9426-3043F1919CD2");
            private const string Owner="KPLN.TEP.Spec20260826";
            public Settings Config {get;private set;}=new Settings();
            public List<Source> Sources {get;private set;}=new List<Source>();
            public List<string> Parameters {get;private set;}=new List<string>();
            public List<string> Categories {get;private set;}=new List<string>();
            public List<string> Phases {get;private set;}=new List<string>();
            public List<string> AreaSchemes {get;private set;}=new List<string>();
            public List<Choice> Templates {get;private set;}=new List<Choice>();
            public Run Last {get;private set;}
            public Detail RequestedDetail {get;set;}
            private Run current;
            private double zero,ground;
            private readonly Dictionary<string,Solid> shapes=new Dictionary<string,Solid>();
            private readonly HashSet<string> notices=new HashSet<string>();
            private readonly Dictionary<string,VolumeSet> volumeSets=new Dictionary<string,VolumeSet>(StringComparer.Ordinal);
            private readonly Dictionary<Element,List<Solid>> rawVolumeSolids=new Dictionary<Element,List<Solid>>();
            private readonly Dictionary<string,List<Solid>> worldVolumeSolids=new Dictionary<string,List<Solid>>(StringComparer.Ordinal);
            private readonly Dictionary<Tuple<Document,bool>,SpatialElementGeometryCalculator> spatialCalculators=new Dictionary<Tuple<Document,bool>,SpatialElementGeometryCalculator>();
            private readonly Dictionary<Tuple<Element,bool>,Solid> localSpatialVolumes=new Dictionary<Tuple<Element,bool>,Solid>();
            private readonly Dictionary<string,Timing> timings=new Dictionary<string,Timing>();
            private readonly List<SlowOperation> slowOperations=new List<SlowOperation>();
            private int volumeCacheHits,spatialCacheHits;
            private int booleanTouchSkips,booleanSplitRecoveries,booleanIntersectionRecoveries,booleanNormalizedRecoveries;
            private readonly HashSet<Correction> appliedCorrections=new HashSet<Correction>();
            private readonly List<Issue> startupIssues=new List<Issue>();
            private bool settingsLoadFailed;
            public Engine(UIApplication application)
            {
                app=application;ui=app.ActiveUIDocument;doc=ui.Document;
                selection=new HashSet<long>(ui.Selection.GetElementIds().Select(IDHelper.ElIdValue));
            }
            public bool IsInitialized {get;private set;}
            private Action<string> reportProgress;
            private bool? createViewsForRun;
            private bool ViewsEnabled {get{return createViewsForRun??Config.CreateViews;}}
            private string progressMessage;
            public string StorageDescription {get{return "В текущем RVT: служебный элемент DataStorage «KPLN.TEP.Spec20260826/settings», Extensible Storage, схема "+StorageId+", поле Payload. Это не параметр проекта. Для записи на диск сохраните RVT. Последний расчёт хранится аналогично в /run.";}}
            private bool catalogsReady;
            private readonly Dictionary<Element,Dictionary<string,IList<Parameter>>> parameterCache=new Dictionary<Element,Dictionary<string,IList<Parameter>>>();
            private readonly Dictionary<Document,Phase> phaseCache=new Dictionary<Document,Phase>();
            public string PendingContourAction {get;set;}
            public string ContourMetricKey {get;set;} = "Gns";
            public bool ReopenContours {get;set;}
            private const double ArcChordTolerance=.0001/.3048; // 0.1 mm, reported in every run.
            [ThreadStatic] private static Dictionary<Solid,Tuple<LayeredBody,string>> planarBodies;
            [ThreadStatic] private static Action planarCheckpoint;
            private int planarVolumeCuts;
            private const string CompressedStoragePrefix="gzip-base64:v1:";

            private TepMethodology Methodology => TepMethodology.For(Config.Method);
        }
    }
}
