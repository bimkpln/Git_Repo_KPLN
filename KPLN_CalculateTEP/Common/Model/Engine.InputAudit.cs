using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public class InputAuditRow
        {
            public string Source {get;set;}
            public string SourceKey {get;set;}
            public bool Critical {get;set;}
            public bool Required {get;set;} = true;
            public int SortPriority {get {return Critical?0:Problem?(Required?1:2):3;}}
            public string Priority {get {return Critical?"Критическая ошибка":Problem?(Required?"Замечания к обязательным данным":"Замечания к необязательным данным"):"Без замечаний";}}
            public string ElementId {get;set;}
            public string Name {get;set;}
            public string Parameter {get;set;}
            public string Value {get;set;}
            public string State {get;set;}
            public bool Problem {get;set;}
            public string Info {get {return string.IsNullOrWhiteSpace(Value)?State:string.IsNullOrWhiteSpace(State)||Value==State?Value:Value+" - "+State;}}
            public string TechnicalCode {get;set;}
            public string ProblemGroupTitle {get
            {
                if(State=="Параметр отсутствует"||(State??"").StartsWith("Отсутствует;"))return "Отсутствует параметр «"+Parameter+"»";
                if((State??"").StartsWith("Не заполнено"))return "Не заполнен параметр «"+Parameter+"»";
                if(!Problem&&!Critical)return "Заполнено / проверено: "+Parameter;
                if(TechnicalCode=="AUTO_SHAFT_BUILDING")return "Не определён корпус шахт";
                if(TechnicalCode=="ELEMENT_BUILDING_UNRESOLVED")return "Не определён корпус конструкций";
                if(TechnicalCode=="AUTO_SHAFT_CONTINUITY")return "Неоднозначные границы шахт на одном этаже";
                if(TechnicalCode=="AUTO_SHAFT_UNCONFIRMED"||TechnicalCode=="AUTO_SHAFT_REGION")return "Границы шахт не подтверждены геометрией";
                if(TechnicalCode=="ROOMS_EMPTY")return "Нет объектов для расчёта площадей";
                if(Critical&&(Parameter=="Площадь"||TechnicalCode=="ZERO_AREA"))return "Нулевая площадь";
                return Parameter+": "+State;
            }}
            public string ProblemGroupKey {get {return SortPriority+"|"+ProblemGroupTitle;}}
            public string ElementGroupKey {get {string source=SourceKey??Source??"",id=ElementId??"-";return source.Length+":"+source+id.Length+":"+id+(id=="-"||id.Length==0?Name??"":"");}}
            public string ElementGroupTitle {get {return (string.IsNullOrEmpty(ElementId)||ElementId=="-"?"Общая проверка":"Элемент "+ElementId)+
                (string.IsNullOrEmpty(Name)?"":" | "+Name)+(string.IsNullOrEmpty(Source)?"":" | "+Source)+" | "+Priority;}}
        }
        public class InputAudit
        {
            public List<InputAuditRow> Rows {get;set;} = new List<InputAuditRow>();
            public BuildingParameterReview Building {get;set;}
            public List<Issue> Issues {get;set;} = new List<Issue>();
        }
        public partial class Engine
        {
            public static IEnumerable<InputAuditRow> OrderAuditRows(IEnumerable<InputAuditRow> rows)
            {return rows.OrderBy(r=>r.SortPriority).ThenBy(r=>r.Source).ThenBy(r=>r.ElementId).ThenBy(r=>r.Parameter);}
            public static string AuditCheckLabel(string parameterKey,string code)
            {
                var parameter=Settings.DefaultParameters().FirstOrDefault(p=>p.Key==parameterKey);
                if(parameter!=null)return parameter.Title;
                switch(parameterKey)
                {
                    case "$level-top":return "Верх перекрытия цоколя";
                    case "$level-normative":return "Нормативные характеристики этажа";
                }
                switch(code)
                {
                    case "ZERO_AREA":case "SPATIAL_UNBOUNDED":return "Нулевая площадь";
                    case "PREFLIGHT_RULE":return "Проверка правил расчёта";
                    case "PARAM_VALUE":case "PARAM_AMBIGUOUS":case "PARAM_SCAN":return "Чтение и заполнение параметров";
                    case "CLASSIFICATION":case "CLASSIFICATION_DICTIONARY":return "Классификация помещений";
                    case "CATEGORY_FAMILY_EMPTY":return "Семейства категории";
                    case "AUTO_DIMENSION":return "Размеры по геометрии";
                    case "AUTO_REFERENCE":case "ROOM_LEVEL":case "PARKING_LEVEL":return "Расчётный уровень объекта";
                    case "AUTO_STAIR_SOURCE":return "Лестницы и марши";
                    case "HEIGHT_GEOMETRY":return "Высота в свету и наклон потолка";
                    case "AUTO_SHAFTS":case "AUTO_SHAFT_GEOMETRY":case "AUTO_SHAFT_CONTINUITY":case "AUTO_SHAFT_UNCONFIRMED":case "AUTO_SHAFT_REGION":case "AUTO_SHAFT_CONFIRMED":return "Геометрия и непрерывность шахт";
                    case "AUTO_SHAFT_BUILDING":return "Корпус шахт";
                    case "SOURCE_UNLOADED":return "Доступность источника";
                    case "SOURCE_TILTED":return "Наклон связанной модели";
                    case "ROOMS_EMPTY":return "Наличие пригодных помещений";
                    case "ROOM_AREA_SETTINGS":return "Границы расчёта площади помещений";
                    case "APARTMENT_PARAMETER":return "Номер квартиры";
                    case "APARTMENT_MULTILEVEL":return "Квартиры на нескольких этажах";
                    case "PARKING_MARK":return "Марка машино-места";
                    case "ZERO_DATUM":return "Ноль здания";
                    case "GROUND_DATUM":return "Уровень земли";
                    case "LEVEL_VALUE":case "LEVEL_GEOMETRY":return "Характеристики этажа";
                    case "LEVEL_TOP_SLAB":return "Верх перекрытия цоколя";
                    case "ELEMENT_BUILDING_UNRESOLVED":case "SINGLE_BUILDING_ASSUMPTION":return "Определение корпуса";
                    case "NO_METRICS":return "Выбор показателей";
                    case "MIXED_PROFILE":return "Профиль смешанного комплекса";
                    case "SETTINGS_MIGRATION":case "SETTINGS_RECOVERY":case "SETTINGS_REFERENCES":return "Сохранённые настройки";
                    case "LINK_AUTOLOAD":return "Загрузка связей";
                    default:return "Проверка исходных данных";
                }
            }
            public void RequestAuditNavigation(InputAuditRow row)
            {
                long id;if(row==null||!long.TryParse(row.ElementId,out id))throw new InvalidOperationException("Эта проверка относится ко всему проекту. У строки нет ID элемента.");
                var matches=Sources.Where(s=>!string.IsNullOrEmpty(row.SourceKey)?s.Key==row.SourceKey:s.Name==row.Source).ToList();
                if(matches.Count!=1||!matches[0].Loaded)throw new InvalidOperationException("Источник элемента недоступен или неоднозначен.");
                var element=matches[0].Document.GetElement(IDHelper.CreateElementId(id));
                if(element==null)throw new InvalidOperationException("Элемент больше не существует в источнике.");
                RequestedDetail=new Detail{SourceKey=matches[0].Key,Element=row.ElementId,UniqueId=element.UniqueId};
            }
            public static string AuditParameterState(bool exists,string value,bool required)
            {
                if(!exists)return required?"Параметр отсутствует":"Отсутствует; не обязателен для этого назначения";
                if(string.IsNullOrWhiteSpace(value))return required?"Не заполнено":"Не заполнено; не обязательно для этого назначения";
                return "Заполнено";
            }
            public void AcceptAuditBuilding(InputAudit review,bool useSingleBuilding)
            {
                singleBuildingAssumption=null;
                if(!useSingleBuilding)return;
                if(review?.Building?.CanAssumeSingle!=true)throw new InvalidOperationException("Для этих источников нельзя подтвердить один корпус: обнаружены разные корпуса или недоступные данные.");
                singleBuildingAssumption=review.Building;
            }
            public InputAudit AuditInputs(Action<string> progress=null)
            {
                ClearSingleBuildingAssumption();return AuditInputsCore(progress);
            }
            public InputAudit RefreshAcceptedInputAudit(InputAudit original,Action<string> progress=null)
            {
                var accepted=singleBuildingAssumption;
                if(accepted==null)return original;
                try{return AuditInputsCore(progress);}
                finally{singleBuildingAssumption=accepted;}
            }
            private InputAudit AuditInputsCore(Action<string> progress)
            {
                PrepareRoomWorkflow();LoadClassificationDictionary();
                parameterCache.Clear();phaseCache.Clear();
                var audit=new InputAudit();var buildings=new List<string>();int buildingErrors=0,index=0;
                foreach(var source in Sources.Where(s=>s.Mode!="exclude"&&s.Loaded&&s.LoadError==null))
                    foreach(var element in source.Elements)
                    {
                        var room=element as Room;
                        if(room!=null&&!IsPlacedRoom(room))continue;
                        bool family=IsClassifiedFamily(element),parking=IsParkingFamily(element);
                        if(room==null&&!family&&!parking)continue;
                        if(!PhaseAccepted(source,element))continue;
                        if(++index%50==0)progress?.Invoke("Отчёт заполнения параметров: объектов "+index);
                        Action<string,string,bool> parameter=(label,name,required)=>
                        {
                            var row=new InputAuditRow{Source=source.Name,SourceKey=source.Key,ElementId=IDHelper.ElIdValue(element.Id).ToString(),Name=element.Name,Parameter=label,Required=required};
                            try
                            {
                                var p=name=="@ParkingMark"?element.get_Parameter(BuiltInParameter.ALL_MODEL_MARK):
                                    name=="@Department"?element.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT):
                                    name=="@Name"?element.get_Parameter(BuiltInParameter.ROOM_NAME):Parameter(element,name);
                                row.Value=Value(element,name);row.State=AuditParameterState(p!=null,row.Value,required);
                                row.Problem=p==null||string.IsNullOrWhiteSpace(row.Value);
                                if(room!=null&&name=="ПОМ_Корпус"&&row.Problem)
                                {var assigned=RoomBuildingValue(room,source);if(!string.IsNullOrWhiteSpace(assigned)){row.Value=assigned;row.State=singleBuildingAssumption!=null?"Корпус принят по подтверждению единого корпуса; параметр модели не изменён":"Корпус задан соответствием в настройках ТЭП; параметр модели не изменён";row.Problem=false;}}
                                if(required&&name=="КВ_Номер"&&(row.Value=="0"||row.Value=="-")){row.State="Не задан номер квартиры (0 или -)";row.Problem=true;}
                            }
                            catch(Exception ex){row.State=ex.Message;row.Problem=true;}
                            audit.Rows.Add(row);
                        };
                            bool apartment=false;
                        if(room!=null||family)
                        {
                            string id=null,role=null;
                            try{id=Mapped(element,"apartment");role=room!=null?DepartmentRole(ClassificationValue(element)):ClassifiedFamilyRole(element);}
                            catch(Exception ex){audit.Rows.Add(new InputAuditRow{Source=source.Name,SourceKey=source.Key,ElementId=IDHelper.ElIdValue(element.Id).ToString(),Name=element.Name,Parameter="Классификация",State=ex.Message,Problem=true});}
                            apartment=!string.IsNullOrWhiteSpace(id)&&id!="0"&&id!="-"||OneOf(role,"heated","auxiliary");
                            parameter("ПОМ_Корпус","ПОМ_Корпус",room!=null);
                            parameter("КВ_Номер","КВ_Номер",apartment);
                            parameter("ПОМ_Секция","ПОМ_Секция",apartment);
                            if(room!=null)parameter("Назначение","@Department",true);
                            if(family&&UsesFamilyParameterArea(element))parameter("Площадь семейства",FamilyAreaParameterName(element),true);
                            if(role=="unknown")audit.Rows.Add(new InputAuditRow{Source=source.Name,SourceKey=source.Key,ElementId=IDHelper.ElIdValue(element.Id).ToString(),Name=element.Name,Parameter="Классификация",State="Не распределено по категориям",Problem=true});
                            var requiredKeys=RequiredParameterKeys(Config,new[]{role});
                            foreach(var map in Config.Parameters.Where(m=>!Settings.IsFixedParameter(m.Key)&&!string.IsNullOrWhiteSpace(m.Name)))
                                parameter(map.Title+" / "+map.Name,map.Name,requiredKeys.Contains(map.Key));
                        }
                        if(parking)parameter("Марка машино-места","@ParkingMark",true);
                        if(room!=null)
                        {
                            try{buildings.Add(RoomBuildingValue(room,source));}catch{buildingErrors++;}
                            if(room.Area<=1e-9)audit.Rows.Add(new InputAuditRow{Source=source.Name,SourceKey=source.Key,ElementId=IDHelper.ElIdValue(room.Id).ToString(),Name=room.Name,Parameter="Площадь",Value="0 м²",State="Нулевая площадь. Критическая ошибка; объект будет пропущен",Problem=true,Critical=true});
                        }
                    }
                audit.Building=ReviewBuildingValues(buildings);audit.Building.ReadErrors=buildingErrors;
                audit.Building.UnavailableSources=Sources.Count(s=>s.Mode!="exclude"&&(!s.Loaded||s.LoadError!=null));
                audit.Issues=CheckParameters(progress);
                audit.Rows.AddRange(LevelGeometryAuditRows());
                foreach(var level in Config.Levels.Where(l=>l.Include&&l.Kind=="basement"&&l.Above=="auto"))
                {
                    var review=TopSlabReviewFor(level);bool manual=level.TopSlabMode=="manual";
                    audit.Rows.Add(new InputAuditRow{Source=level.Source,ElementId="-",Name=level.Name,Parameter="Верх перекрытия цоколя",
                        Value=level.TopSlabSummary,State=manual?"Ручная абсолютная отметка; геометрия не подтверждает это значение":review.Selected?.Elements??review.Error,
                        Problem=manual||review.Selected==null});
                }
                var names=new Dictionary<string,string>();
                foreach(var issue in audit.Issues.Where(i=>(i.Severity!="Информация"||OneOf(i.Code,"AUTO_DIMENSION","AUTO_REFERENCE","AUTO_STAIR_SOURCE","HEIGHT_GEOMETRY"))&&i.Code!="SPATIAL_UNBOUNDED"))
                foreach(var elementId in issue.ElementIds??new List<string>{issue.Element})
                {
                    var source=Sources.Count(s=>s.Name==issue.Source)==1?Sources.Single(s=>s.Name==issue.Source):null;
                    string name=null;long id;
                    if(source?.Loaded==true&&long.TryParse(elementId,out id))
                    {
                        string key=source.Key+"/"+elementId;
                        if(!names.TryGetValue(key,out name))
                        {name=source.Document.GetElement(IDHelper.CreateElementId(id))?.Name;names[key]=name;}
                    }
                    bool missingBuilding=issue.Code=="CLASSIFICATION"&&(issue.Message==MissingBuildingMessage(null)||issue.Message==MissingBuildingMessage(""));
                    audit.Rows.Add(new InputAuditRow{Source=issue.Source,SourceKey=source?.Key,Name=name,ElementId=string.IsNullOrEmpty(elementId)?"-":elementId,Parameter=missingBuilding?"ПОМ_Корпус":AuditCheckLabel(issue.ParameterKey,issue.Code),TechnicalCode=issue.Code,
                        State=missingBuilding?(issue.Message==MissingBuildingMessage(null)?"Параметр отсутствует":"Не заполнено"):issue.Message,Problem=issue.Severity!="Информация",Required=issue.Severity=="Ошибка",Critical=OneOf(issue.Code,"ZERO_AREA","SPATIAL_UNBOUNDED","SOURCE_UNLOADED","SOURCE_TILTED","ROOMS_EMPTY")});
                }
                foreach(var group in audit.Rows.GroupBy(r=>r.ElementGroupKey))
                {
                    string name=group.Select(r=>r.Name).FirstOrDefault(n=>!string.IsNullOrEmpty(n));
                    foreach(var row in group.Where(r=>string.IsNullOrEmpty(r.Name)))row.Name=name;
                }
                audit.Rows=DistinctAuditRows(audit.Rows);
                return audit;
            }
            public static List<InputAuditRow> DistinctAuditRows(IEnumerable<InputAuditRow> rows)
            {
                // One identical check per element even when several indicators report it.
                return OrderAuditRows(rows.GroupBy(r=>new{r.ElementGroupKey,r.Parameter,r.Value,r.State,r.SortPriority}).Select(g=>g.First())).ToList();
            }
        }
    }
}
