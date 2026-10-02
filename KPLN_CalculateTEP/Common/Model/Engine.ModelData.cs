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
            public List<Issue> CheckParameters(Action<string> progress=null)
            {
                PrepareRoomWorkflow(); Normalize(Config); parameterCache.Clear(); phaseCache.Clear(); reportProgress=progress;
                var saved=current; current=new Run(); current.Issues.AddRange(startupIssues); notices.Clear(); invalidRoomFloors.Clear();
                try
                {
                    ValidateDatums(); PreflightRecords();
                    return current.Issues.ToList();
                }
                finally {current=saved; reportProgress=null; parameterCache.Clear();}
            }
            private void ValidateDatums()
            {
                zero=double.NaN;ground=double.NaN;
                if(Config.Metrics.Any(m=>m.Enabled&&MetricNeedsZero(m.Key)))
                    try {zero=Datum(Config.ZeroMode,Config.ZeroValue,Config.ZeroLevel,Config.ZeroParameter,false);}
                    catch(OperationCanceledException){throw;}
                    catch(Exception ex){current.Issue("ZERO_DATUM","Ошибка",ex.Message);}
                if(Config.Metrics.Any(m=>m.Enabled&&MetricNeedsGround(m.Key)))
                    try {ground=Datum(Config.GroundMode,Config.GroundValue,Config.GroundLevel,Config.GroundParameter,true);}
                    catch(OperationCanceledException){throw;}
                    catch(Exception ex){current.Issue("GROUND_DATUM","Ошибка",ex.Message);}
            }
            private static bool Eq(string a,string b){return string.Equals((a??"").Trim(),(b??"").Trim(),StringComparison.OrdinalIgnoreCase);}
            private IList<Parameter> NamedParameters(Element e,string name)
            {
                Dictionary<string,IList<Parameter>> names;if(!parameterCache.TryGetValue(e,out names))parameterCache[e]=names=new Dictionary<string,IList<Parameter>>(StringComparer.Ordinal);
                IList<Parameter> result;if(!names.TryGetValue(name,out result))names[name]=result=e.GetParameters(name);return result;
            }
            private Parameter Parameter(Element e,string name)
            {
                if(string.IsNullOrWhiteSpace(name)||name.StartsWith("@"))return null;
                var found=NamedParameters(e,name);
                if(found.Count==0){var type=e.Document.GetElement(e.GetTypeId());if(type!=null)found=NamedParameters(type,name);}
                if(found.Count>1)throw new InvalidOperationException("Несколько параметров с именем «"+name+"».");
                return found.FirstOrDefault();
            }
            private string Value(Element e,string name)
            {
                if(string.IsNullOrWhiteSpace(name))return null;
                if(name=="@ParkingMark")return e.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)?.AsString();
                if(name=="@Department")return e.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT)?.AsString()??(e is Room?"":null);
                if(name=="@Name")return e.Name;
                if(name=="@Category")return e.Category==null?null:e.Category.Name;
                if(name=="@Type")return e.Document.GetElement(e.GetTypeId())?.Name;
                if(name=="@Family")return (e as FamilyInstance)?.Symbol.FamilyName;
                if(name=="@Workset")return e.Document.GetWorksetTable().GetWorkset(e.WorksetId)?.Name;
                var p=Parameter(e,name);if(p==null)return null;if(!p.HasValue)return "";
                if(p.StorageType==StorageType.String)return p.AsString()??"";
                if(p.StorageType==StorageType.Integer)return p.AsInteger().ToString(CultureInfo.InvariantCulture);
                if(p.StorageType==StorageType.Double)return p.AsValueString()??p.AsDouble().ToString("R",CultureInfo.InvariantCulture);
                return p.AsValueString()??IDHelper.ElIdValue(p.AsElementId()).ToString();
            }
            private string Mapped(Element e,string key){return Value(e,Config.Parameter(key));}
            private double? Number(Element e,string key,bool length=false)
            {
                string name=Config.Parameter(key);var p=Parameter(e,name);if(p==null||!p.HasValue)return null;
                if(p.StorageType==StorageType.Double)
                {
                    if(length&&!IDHelper.IsLength(p))throw new InvalidOperationException("Параметр «"+name+"» должен иметь тип Длина, либо текст в метрах.");
                    if(!length&&!IDHelper.IsNumber(p)&&!IDHelper.IsAngle(p))throw new InvalidOperationException("Параметр «"+name+"» должен быть безразмерным числом или углом.");
                    return p.AsDouble()*(length?.3048:IDHelper.IsAngle(p)?180/Math.PI:1);
                }
                if(length&&p.StorageType!=StorageType.String)throw new InvalidOperationException("Параметр «"+name+"» должен иметь тип Длина либо содержать текстовое число в метрах.");
                double n;return TryNumber(Value(e,name),out n)?(double?)n:null;
            }
            public static bool TryNumber(string value,out double result)
            {return double.TryParse((value??"").Trim().Replace(",","."),NumberStyles.Float,CultureInfo.InvariantCulture,out result)&&!double.IsNaN(result)&&!double.IsInfinity(result);}
            private static double RequiredNumber(string value,string title)
            {double n;if(!TryNumber(value,out n))throw new InvalidOperationException("Укажите число: "+title+".");return n;}
            private bool Flag(Element e,string key)
            {var v=Mapped(e,key);if(string.IsNullOrWhiteSpace(v))return false;if(Eq(v,"1")||Eq(v,"true")||Eq(v,"да"))return true;if(Eq(v,"0")||Eq(v,"false")||Eq(v,"нет"))return false;throw new InvalidOperationException("Признак «"+Config.Parameter(key)+"» должен содержать да/нет или 1/0.");}
            private bool Matches(Element e,Rule rule)
            {
                if(!rule.Enabled||(!string.IsNullOrWhiteSpace(rule.Category)&&!Eq(rule.Category,e.Category?.Name)))return false;
                string v=Value(e,rule.Parameter);if(v==null)return false;
                if(rule.Operation=="empty")return v.Trim().Length==0;if(rule.Operation=="filled")return v.Trim().Length>0;
                if(rule.Operation=="equals")return Eq(v,rule.Value);
                if(rule.Operation=="contains")return !string.IsNullOrEmpty(rule.Value)&&v.IndexOf(rule.Value,StringComparison.OrdinalIgnoreCase)>=0;
                throw new InvalidOperationException("Неизвестное условие правила: "+rule.Operation);
            }
            private static string RoleKey(string input)
            {return RoleLabels.FirstOrDefault(x=>Eq(x.Key,input)||Eq(x.Value,input)).Key;}
            private bool PhaseAccepted(Source source,Element e)
            {
                if(e.DesignOption!=null&&!e.DesignOption.IsPrimary)return false;
                Phase phase;if(!phaseCache.TryGetValue(source.Document,out phase))
                {var phases=source.Document.Phases.Cast<Phase>().ToList();phase=string.IsNullOrWhiteSpace(Config.Phase)?phases.LastOrDefault():phases.FirstOrDefault(p=>Eq(p.Name,Config.Phase));phaseCache[source.Document]=phase;}
                if(phase==null)throw new InvalidOperationException("В источнике отсутствует выбранная стадия «"+Config.Phase+"».");
                if(e is SpatialElement)
                {var parameter=e.get_Parameter(BuiltInParameter.ROOM_PHASE_ID);if(parameter!=null&&parameter.StorageType==StorageType.ElementId&&parameter.AsElementId()!=ElementId.InvalidElementId)return parameter.AsElementId()==phase.Id;}
                if(!e.HasPhases())return true;
                var status=e.GetPhaseStatus(phase.Id);return status==ElementOnPhaseStatus.New||status==ElementOnPhaseStatus.Existing;
            }
            private string RecordParameter(Source source,Element element,string key)
            {
                try{return Mapped(element,key)??"";}
                catch(OperationCanceledException){throw;}
                catch(Exception ex){current.Issues.Add(new Issue{Code="PARAM_AMBIGUOUS",Severity="Ошибка",Message=ex.Message,ParameterKey=key,
                    InputKind=element is Room?"room":"",Source=source.Name,Element=IDHelper.ElIdValue(element.Id).ToString()});return "";}
            }
            private Record MakeRecord(Source source,Element e)
            {
                var level=e.Document.GetElement(e.LevelId) as Level;
                if(level==null&&e is SpatialElement)level=(e as SpatialElement).Level;
                string levelName=Mapped(e,"level");
                if(!string.IsNullOrWhiteSpace(levelName))
                {
                    var parameter=Parameter(e,Config.Parameter("level"));
                    if(parameter!=null&&parameter.StorageType==StorageType.ElementId)level=e.Document.GetElement(parameter.AsElementId()) as Level;
                    else level=new FilteredElementCollector(e.Document).OfClass(typeof(Level)).Cast<Level>().SingleOrDefault(l=>Eq(l.Name,levelName));
                    if(level==null)throw new InvalidOperationException("Параметр расчётного уровня не соответствует уровню модели: "+levelName);
                }
                var setting=level==null?null:Config.Levels.FirstOrDefault(x=>x.Key==source.Key+"/"+level.UniqueId);
                var r=new Record{Source=source,Element=e,Level=setting,Section=e is Room?Mapped(e,"section")??"":"",Apartment=e is Room?RecordParameter(source,e,"apartment"):"",Vertical=RecordParameter(source,e,"vertical"),Role="unknown",Part="auto"};
                r.Apartment=(r.Apartment??"").Trim(); if(r.Apartment=="0"||r.Apartment=="-")r.Apartment="";
                string value=Config.Grouping=="links"?source.Name:Config.Grouping=="worksets"?Value(e,"@Workset"):
                    Config.Grouping=="selection"?Config.ManualBuilding:e is Room?Mapped(e,"building"):null;
                var maps=Config.Buildings.Where(x=>(x.Source=="*"||Eq(x.Source,source.Key)||Eq(x.Source,source.Name))&&(x.MatchValue=="*"||Eq(x.MatchValue,value))).ToList();
                if(maps.Count>1)throw new InvalidOperationException("Объект соответствует нескольким строкам назначения корпуса.");
                var map=maps.FirstOrDefault();r.Building=map?.Building??value;
                if(e is Room&&string.IsNullOrWhiteSpace(r.Building))throw new InvalidOperationException("У размещённого помещения не заполнен обязательный параметр «ПОМ_Корпус».");
                r.BuildingIncluded=map?.Include??true;
                if(level!=null)
                {
                    var overrides=Config.Levels.Where(l=>l.Key==source.Key+"/"+level.UniqueId&&(!string.IsNullOrWhiteSpace(l.Building)||!string.IsNullOrWhiteSpace(l.Section))&&
                        (string.IsNullOrWhiteSpace(l.Building)||Eq(l.Building,r.Building))&&(string.IsNullOrWhiteSpace(l.Section)||Eq(l.Section,r.Section))).ToList();
                    if(overrides.Count>1)throw new InvalidOperationException("Несколько настроек уровня соответствуют одному корпусу и секции.");
                    if(overrides.Count==1)r.Level=overrides[0];
                }
                r.Profile=Config.Profile=="by-source"?source.Profile:Config.Profile=="by-building"?map?.Profile:Config.Profile;
                if(r.Profile=="high-mixed"&&map!=null&&OneOf(map.Profile,"high-residential","high-public"))r.Profile=map.Profile;
                string mappedProfile=Mapped(e,"profile");if(!string.IsNullOrWhiteSpace(mappedProfile))
                {var p=Choices("profile").FirstOrDefault(x=>Eq(x.Key,mappedProfile)||Eq(x.Label,mappedProfile));if(p==null)throw new InvalidOperationException("Неизвестный профиль здания: "+mappedProfile);r.Profile=p.Key;}
                if(string.IsNullOrWhiteSpace(r.Profile)||r.Profile.StartsWith("by-"))throw new InvalidOperationException("Задайте профиль для корпуса / источника.");
                r.BuildingClass=map?.Class??(r.Profile.Contains("public")?"nonresidential":"residential");
                if(r.Profile=="high-mixed")throw new InvalidOperationException("Для многофункционального высотного здания назначьте элементам профиль функциональной части через параметр типа здания (high-residential / high-public). Требования СП 160 не подменяются одним общим правилом.");
                if(Config.Profile=="high-mixed")Notice("MIXED_PROFILE","Предупреждение","Для смешанного комплекса применён заданный профиль функциональной части. Требуется экспертная сверка состава частей по СП 160.",r);
                if(e is HostObject||e is Wall)r.Role="structure";
                if(e.Category!=null&&IDHelper.ElIdValue(e.Category.Id)==(long)BuiltInCategory.OST_StructuralFoundation)r.Role="structure";
                if(IsParkingFamily(e))r.Role="parking";
                if(e is Opening)r.Role="opening";
                if(e is Room)r.Role=DepartmentRole(Value(e,"@Department"));
                if(e is Room)r.Part=RoomPart(Value(e,"@Department"));
                ApplyRules(r,"all");
                // Classification is validated after metric-specific rules in the preflight pass.
                return r;
            }
            private void ApplyRules(Record r,string metric)
            {
                var applicable=Config.Rules.Where(x=>r.Element is Room&&x.Action=="classify"&&x.Metric==metric&&Matches(r.Element,x)).ToList();
                var classifications=applicable.Where(x=>x.Action=="classify").ToList();
                if(classifications.Select(x=>x.Role+"|"+x.Part+"|"+x.Coefficient).Distinct().Count()>1)
                    throw new InvalidOperationException("Конфликт классифицирующих правил для "+metric+".");
                if(applicable.Any(x=>x.Action=="include")&&applicable.Any(x=>x.Action=="exclude"))throw new InvalidOperationException("Одновременное включение и исключение правилами.");
                foreach(var rule in applicable)
                {
                    if(rule.Action=="classify"){r.Role=rule.Role;r.Part=rule.Part;r.Factor=rule.Coefficient;if(rule.Coefficient.HasValue)r.Manual=true;}
                    if(rule.Action=="include"){r.Override=true;r.Manual=true;}
                    if(rule.Action=="exclude"){r.Override=false;r.Manual=true;}
                }
                foreach(var c in Config.Corrections.Where(c=>(c.Metric==metric||(metric=="all"&&string.IsNullOrEmpty(c.Metric)))&&c.Action!="delta"&&
                    (c.Source==r.Source.Key||Eq(c.Source,r.Source.Name))&&(c.Element==IDHelper.ElIdValue(r.Element.Id).ToString()||c.Element==r.Element.UniqueId)))
                {
                    if(string.IsNullOrWhiteSpace(c.Reason))throw new InvalidOperationException("Для ручной корректировки необходима причина.");
                    if(c.Action=="include"||c.Action=="exclude"){c.Previous=r.Override?.ToString()??"По нормативу";r.Override=c.Action=="include";}
                    else if(c.Action=="classify"){c.Previous=r.Role;r.Role=RoleKey(c.Value)??throw new InvalidOperationException("Неизвестная роль корректировки.");}
                    else if(c.Action=="building"){c.Previous=r.Building;if(string.IsNullOrWhiteSpace(c.Value))throw new InvalidOperationException("Пустой корпус в корректировке.");r.Building=c.Value;}
                    r.Manual=true;
                    appliedCorrections.Add(c);
                }
            }
            private double Datum(string mode,string value,string levelKey,string parameter,bool isGround)
            {
                if(mode=="manual")return RequiredNumber(value,isGround?"отметка земли, м":"отметка 0.000, м")/.3048;
                if(mode=="level")
                {var l=Config.Levels.FirstOrDefault(x=>x.Key==levelKey);if(l==null)throw new InvalidOperationException(isGround?"Не выбран уровень земли.":"Выберите уровень первого этажа для нуля здания во вкладке «Квартиры, этажи и отметки».");return l.Elevation;}
                if(mode=="parameter")
                {
                    var p=Parameter(doc.ProjectInformation,parameter);if(p==null||!p.HasValue)throw new InvalidOperationException("Не заполнен параметр отметки в сведениях о проекте: "+parameter);
                    if(p.StorageType==StorageType.Double&&IDHelper.IsLength(p))return p.AsDouble();
                    return RequiredNumber(Value(doc.ProjectInformation,parameter),parameter)/.3048;
                }
                if(mode=="terrain"&&isGround)
                {
                    var elevations=new List<double>();
                    foreach(var s in Sources.Where(x=>x.Loaded&&x.Mode=="include"))foreach(var e in s.Elements)
                    {
                        long category=IDHelper.ElIdValue(e.Category.Id);
                        // Toposolid is discovered by its runtime name to keep the same source compatible with Revit 2020.
                        if(category!=(long)BuiltInCategory.OST_Topography&&e.GetType().Name!="Toposolid")continue;
                        if(selection.Count>0&&s.RootLink==null&&!selection.Contains(IDHelper.ElIdValue(e.Id)))continue;
                        var top=e as TopographySurface;
                        if(top!=null)elevations.AddRange(top.GetPoints().Select(p=>s.Transform.OfPoint(p).Z));
                        else foreach(var solid in Solids(e))foreach(Face face in solid.Faces)
                            if(face.ComputeNormal(new UV(.5,.5)).Z>0.1)elevations.AddRange(face.Triangulate().Vertices.Select(p=>s.Transform.OfPoint(p).Z));
                    }
                    if(elevations.Count==0)throw new InvalidOperationException("Не найдена поверхность земли. Выберите топографию до запуска или задайте отметку вручную.");
                    if(elevations.Max()-elevations.Min()>0.01/.3048)
                        throw new InvalidOperationException("Планировочная поверхность имеет уклон. Единую отметку земли нельзя определить по всему участку: задайте отметку вручную, а наземность этажей уточните по секциям.");
                    return elevations.Average();
                }
                throw new InvalidOperationException("Неизвестный способ определения отметки.");
            }

            private static bool IsParkingFamily(Element e)
            {return e is FamilyInstance && e.Category!=null && IDHelper.ElIdValue(e.Category.Id)==(long)BuiltInCategory.OST_Parking;}
            private static bool IsPlacedRoom(Room room)
            {return room!=null&&room.Location is LocationPoint;}

            // A room parameter must never be required on walls, slabs or parking families.
            // Infer ownership only when all placed rooms identify one building; never guess between buildings.
            private void AssignNonRoomBuildings(List<Record> records)
            {
                if(Config.Grouping!="parameter"||!records.Any(r=>!(r.Element is Room)&&string.IsNullOrWhiteSpace(r.Building)))return;
                var owners=new Dictionary<string,List<string>>();
                foreach(var source in Sources.Where(s=>s.Loaded&&s.LoadError==null&&s.Mode!="exclude"))
                {
                    var values=new List<string>();int scanned=0;
                    foreach(var room in source.Elements.OfType<Room>().Where(IsPlacedRoom))
                    {
                        if(++scanned%100==0)Progress("Определение корпусов по помещениям: "+source.Name+" - "+scanned);
                        try{if(PhaseAccepted(source,room))values.Add(Mapped(room,"building"));}
                        catch(OperationCanceledException){throw;}
                        catch{values.Add(null);}
                    }
                    owners[source.Key]=values;
                }
                foreach(var group in records.Where(r=>!(r.Element is Room)&&string.IsNullOrWhiteSpace(r.Building)).GroupBy(r=>r.Source.Key).ToList())
                {
                    var source=group.First().Source;
                    List<string> values;owners.TryGetValue(source.Key,out values);
                    // Geometry-only links may belong to the sole building in the included model.
                    if(values==null||values.Count==0)
                        values=Sources.Where(s=>s.Mode==source.Mode&&owners.ContainsKey(s.Key)).SelectMany(s=>owners[s.Key]).ToList();
                    string building=UniqueRoomBuilding(values);
                    if(building!=null){foreach(var record in group)record.Building=building;continue;}
                    var affected=Config.Metrics.Where(m=>m.Enabled&&
                        (m.Key=="ParkingCount"?group.Any(r=>IsParkingFamily(r.Element)):
                        group.Any(r=>!IsParkingFamily(r.Element))&&(m.Key.StartsWith("Volume")||OneOf(m.Key,"Footprint","Storeys","Floors")))).ToList();
                    foreach(var metric in affected)
                        current.Issue("ELEMENT_BUILDING_UNRESOLVED","Ошибка",
                            "Не удалось однозначно определить корпус элементов по размещённым помещениям: помещений нет, корпус не заполнен или указано несколько корпусов. Элементов: "+group.Count()+".",
                            metric.Key,source.Name,action:"Параметр ПОМ_Корпус требуется только у помещений. Для элементов нескольких корпусов необходимо пространственное распределение; автоматическое назначение произвольному корпусу не выполняется.");
                }
                records.RemoveAll(r=>!(r.Element is Room)&&string.IsNullOrWhiteSpace(r.Building));
            }
            public static string UniqueRoomBuilding(IEnumerable<string> values)
            {
                var names=(values??Enumerable.Empty<string>()).Select(v=>(v??"").Trim()).ToList();
                if(names.Count==0||names.Any(string.IsNullOrWhiteSpace))return null;
                var distinct=names.Distinct(StringComparer.OrdinalIgnoreCase).Take(2).ToList();
                return distinct.Count==1?distinct[0]:null;
            }
            private List<Record> Collect()
            {
                var records=new List<Record>();
                foreach(var source in Sources.Where(s=>s.Mode!="exclude"))
                {
                    if(!source.Loaded||source.LoadError!=null){current.Issue("SOURCE_UNLOADED","Ошибка",source.LoadError??("Связь «"+source.Name+"» не загружена."),source:source.Name,action:"Проверьте доступность файла и рабочий набор в управлении связями. Перед повторной проверкой или расчётом плагин предложит загрузить незагруженные связи.");continue;}
                    if(Math.Abs(source.Transform.BasisZ.DotProduct(XYZ.BasisZ)-1)>1e-8)
                    {Notice("SOURCE_TILTED","Ошибка","Наклонённая связь не поддерживает поэтажную классификацию: "+source.Name);continue;}
                    var candidates=new List<Record>();
                    int scanned=0,unplaced=0;
                    foreach(var e in source.Elements)
                    {
                        if(++scanned%100==0)Progress("Классификация: "+source.Name+" - "+scanned+" / "+source.Elements.Count);
                        if(e is SpatialElement && !(e is Room))continue;
                        if(e is Room&&!IsPlacedRoom((Room)e)){unplaced++;continue;}
                        if(!(e is SpatialElement)&&!(e is HostObject)&&!(e is FamilyInstance)&&!(e is Opening)&&!(e is DirectShape)&&!(e is Ceiling))continue;
                        if(Config.Grouping=="selection"&&!selection.Contains(IDHelper.ElIdValue(source.RootLink??e.Id)))continue;
                        try
                        {
                            if(!PhaseAccepted(source,e))continue;
                            bool parking=IsParkingFamily(e);
                            bool foundation=e.Category!=null&&IDHelper.ElIdValue(e.Category.Id)==(long)BuiltInCategory.OST_StructuralFoundation;
                            bool configured=Config.Rules.Any(rule=>rule.Enabled&&(rule.Metric=="all"||Config.Metrics.Any(m=>m.Enabled&&m.Key==rule.Metric))&&Matches(e,rule));
                            bool hasFunction=!string.IsNullOrWhiteSpace(Mapped(e,"function"));
                            bool construction=Config.Metrics.Any(m=>m.Enabled&&(m.Key.StartsWith("Volume")||OneOf(m.Key,"Footprint","Storeys","Floors")));
                            if(!(e is SpatialElement)&&!(e is Opening)&&!parking&&!hasFunction&&!configured&&!(construction&&(e is HostObject||foundation)))continue;
                            var spatial=e as SpatialElement;
                            if(spatial!=null&&SpatialArea(spatial)<=1e-9){MarkInvalidRoomFloor(source,e);current.Issue("SPATIAL_UNBOUNDED","Ошибка","Размещённое помещение имеет нулевую площадь. Проверьте замкнутость границ и наличие нескольких помещений в одном контуре.",source:source.Name,element:IDHelper.ElIdValue(e.Id).ToString());continue;}
                            var record=MakeRecord(source,e);if(record.BuildingIncluded)candidates.Add(record);
                        }
                        catch(System.OperationCanceledException){throw;}
                    catch(Exception ex)
                        {
                            MarkInvalidRoomFloor(source,e);
                            var affected=(e is HostObject)?Config.Metrics.Where(m=>m.Enabled&&(m.Key.StartsWith("Volume")||OneOf(m.Key,"Footprint","Storeys","Floors"))).Select(m=>m.Key).ToList():new List<string>{""};
                            foreach(var metric in affected)current.Issues.Add(new Issue{Code="CLASSIFICATION",Severity=source.Mode=="reference"?"Предупреждение":"Ошибка",Message=ex.Message,Metric=metric,Source=source.Name,Element=IDHelper.ElIdValue(e.Id).ToString(),InputKind=e is Room?"room":IsParkingFamily(e)?"parking":""});
                        }
                    }
                    if(unplaced>0)current.Issue("ROOMS_UNPLACED_SKIPPED","Информация","Неразмещённые помещения исключены из расчёта и проверки параметров: "+unplaced+".",source:source.Name,action:"Действия не требуются: рассчитываются только размещённые помещения.");
                    records.AddRange(candidates);
                }
                AssignNonRoomBuildings(records);
                if(Config.Grouping=="selection"&&selection.Count==0)Notice("SELECTION_EMPTY","Ошибка","Ручной режим требует выбора объектов или экземпляров связей до запуска команды.");
                return records;
            }
            private List<Record> Representations(List<Record> records,Metric metric)
            {
                    var selectedSpatial=new HashSet<Record>();
                    foreach(var original in records.Where(r=>r.Element is SpatialElement).GroupBy(r=>r.Source.Key+"|"+r.Level?.Key))
                    {
                        string scheme=string.IsNullOrWhiteSpace(metric.AreaScheme)?Config.AreaScheme:metric.AreaScheme;
                        var group=original.Where(r=>!(r.Element is Area)||string.IsNullOrWhiteSpace(scheme)||Eq(((Area)r.Element).AreaScheme.Name,scheme)).ToList();
                        string kind=string.IsNullOrWhiteSpace(metric.Geometry)?Config.Geometry:metric.Geometry;
                        if(metric.ContourMode=="current"&&(metric.Key.StartsWith("Volume")||metric.Key=="Footprint")||metric.ContourMode!="current"&&kind=="auto")
                            kind=group.Any(r=>r.Element is Room)?"rooms":group.Any(r=>r.Element is Space)?"spaces":kind;
                        if(kind=="auto")kind=group.Any(r=>r.Element is Area)?"areas":group.Any(r=>r.Element is Room)?"rooms":"spaces";
                        var chosen=group.Where(r=>kind=="areas"?r.Element is Area:kind=="rooms"?r.Element is Room:r.Element is Space).ToList();
                        if(kind=="areas"&&chosen.Select(r=>((Area)r.Element).AreaScheme.Id).Distinct().Count()>1)
                        {Notice("AREA_SCHEME","Ошибка","На уровне несколько схем зон. Выберите одну расчётную схему.",group.First(),metric.Key);continue;}
                        if(chosen.Count==0&&group.Count>0)Notice("REPRESENTATION_MISSING","Ошибка","На уровне отсутствует выбранный источник площадей: "+kind,group.First(),metric.Key);
                        foreach(var r in chosen)selectedSpatial.Add(r);
                        if(chosen.Count<original.Count())Notice("REPRESENTATION","Информация","На уровне используется один источник площадей: "+kind+". Параллельные Rooms/Spaces/Areas не суммируются.",original.First(),metric.Key);
                    }
                    return records.Where(r=>!(r.Element is SpatialElement)||selectedSpatial.Contains(r)).ToList();
            }
            private static double SpatialArea(SpatialElement e)
            {return e is Room?((Room)e).Area:e is Space?((Space)e).Area:e is Area?((Area)e).Area:0;}
        }
    }
}
