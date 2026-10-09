using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public partial class Engine
        {
            private bool automaticAreaInputs;
            private readonly Dictionary<string,string> areaBuildFailures=new Dictionary<string,string>();

            private List<Tuple<Record,PlanarRegion>> PersistFootprintInputs(List<Tuple<Record,PlanarRegion>> inputs)
            {
                if(!automaticAreaInputs)return inputs;
                var prepared=new List<Tuple<Record,PlanarRegion>>();
                var keys=new HashSet<string>();
                automaticAreaInputs=false;
                try
                {
                    foreach(var group in inputs.GroupBy(i=>i.Item1.Key))
                    {
                        var record=group.First().Item1;
                        try
                        {
                            var shape=group.Aggregate(PlanarRegion.Empty,(sum,item)=>RegionUnion(sum,item.Item2));
                            var persisted=ReviewRegion("footprint/"+record.Key,record,"Застройка - проекция объекта",ground,()=>shape);
                            prepared.Add(Tuple.Create(record,persisted));
                            keys.Add(Config.Method+"/"+(Config.Phase??"")+"/footprint/"+record.Key);
                        }
                        catch(OperationCanceledException){throw;}
                        catch(Exception ex){Notice("FOOTPRINT_AREA_INPUT","Ошибка",ex.Message,record,"Footprint");}
                    }
                    var pending=ReviewInputs.Where(i=>keys.Contains(i.Key)&&!reviewBindings.Any(b=>b.Key==i.Key)).ToList();
                    try{CreateEditableInputs(reportProgress,pending);}
                    catch(OperationCanceledException){throw;}
                    catch(Exception ex){foreach(var input in pending)input.Status="Не построено: "+ex.Message;}
                    foreach(var input in pending.Where(i=>!reviewBindings.Any(b=>b.Key==i.Key)))areaBuildFailures[input.Key]=input.Status;
                }
                finally{automaticAreaInputs=true;}
                var result=new List<Tuple<Record,PlanarRegion>>();
                foreach(var item in prepared)
                {
                    try{result.Add(Tuple.Create(item.Item1,ReviewRegion("footprint/"+item.Item1.Key,item.Item1,"Застройка - проекция объекта",ground,()=>item.Item2)));}
                    catch(OperationCanceledException){throw;}
                    catch(Exception ex){Notice("FOOTPRINT_AREA_INPUT","Ошибка",ex.Message,item.Item1,"Footprint");}
                }
                return result;
            }

            private void PrepareAutomaticAreaInputs(List<Record> records)
            {
                automaticAreaInputs=false;areaBuildFailures.Clear();OverlayIssues.Clear();
                var rooms=records.Where(r=>IsAreaInput(r.Element)&&r.Level!=null&&r.Level.Include).ToList();
                var known=new HashSet<string>(rooms.Select(r=>r.Key));
                // Even an unclassified room must be visible. This does not authorize its inclusion in a metric.
                foreach(var source in Sources.Where(s=>s.Loaded&&s.LoadError==null&&s.Mode=="include"))
                    foreach(var room in source.Elements.OfType<Room>().Where(r=>IsPlacedRoom(r)&&r.Area>1e-9&&PhaseAccepted(source,r)))
                    {
                        if(known.Contains(source.Key+"/"+room.UniqueId))continue;
                        var level=room.Level;var setting=level==null?null:Config.Levels.FirstOrDefault(l=>l.Key==source.Key+"/"+level.UniqueId);
                        if(setting==null||!setting.Include)continue;
                        var record=new Record{Source=source,Element=room,Level=setting,Profile=Config.Profile,Role="unknown"};
                        try{record.Role=DepartmentRole(ClassificationValue(room));}catch{ }
                        rooms.Add(record);
                    }
                bool envelope=Config.Metrics.Any(m=>m.Enabled&&(RoomEnvelopeTotal((Indicator)Enum.Parse(typeof(Indicator),m.Key))||RoomEnvelopePart((Indicator)Enum.Parse(typeof(Indicator),m.Key))));
                foreach(var room in rooms)
                {
                    // A scalar input deliberately has no editable planar boundary.
                    if(UsesFamilyParameterArea(room.Element))continue;
                    Progress("Подготовка расчётных 2D-областей: "+room.Level.Name+"; ID "+IDHelper.ElIdValue(room.Element.Id));
                    Action<Action> prepare=action=>
                    {
                        try{action();}
                        catch(OperationCanceledException){throw;}
                        catch(Autodesk.Revit.Exceptions.OperationCanceledException){throw;}
                        catch(Autodesk.Revit.Exceptions.RegenerationFailedException){throw;}
                        catch(Exception ex){current.Issue("AREA_INPUT_PREPARE","Предупреждение",ex.Message,source:room.Source.Name,element:IDHelper.ElIdValue(room.Element.Id).ToString());}
                    };
                    prepare(()=>{RoomNetRegion(room);});
                    if(Config.Metrics.Any(m=>m.Enabled&&OneOf(m.Key,"ApartmentsTotal","ApartmentsHeated")))prepare(()=>{MeasuredRoomRegion(room,"ApartmentsTotal");});
                    if(envelope)prepare(()=>{RoomCentreRegion(room);});
                }
                contourRoomRecords=rooms;
                foreach(var shaft in records.Where(r=>r.AutomaticShaftRegion!=null).Concat(unconfirmedAutomaticShafts))
                {
                    try{AutomaticShaftPlan(shaft);}
                    catch(OperationCanceledException){throw;}
                    catch(Autodesk.Revit.Exceptions.OperationCanceledException){throw;}
                    catch(Autodesk.Revit.Exceptions.RegenerationFailedException){throw;}
                    catch(Exception ex){Notice("AUTO_SHAFT_REGION","Ошибка",ex.Message,shaft);}
                }
                if(envelope)
                    foreach(var floor in rooms.GroupBy(r=>Math.Round(r.Z,6)).Select(g=>g.First()))
                        foreach(var metric in Config.Metrics.Where(m=>m.Enabled&&(RoomEnvelopeTotal((Indicator)Enum.Parse(typeof(Indicator),m.Key))||RoomEnvelopePart((Indicator)Enum.Parse(typeof(Indicator),m.Key)))))
                        {
                            try{RoomFloorRegions(records,floor,(Indicator)Enum.Parse(typeof(Indicator),metric.Key));}
                            catch(OperationCanceledException){throw;}
                            catch(Autodesk.Revit.Exceptions.OperationCanceledException){throw;}
                            catch(Autodesk.Revit.Exceptions.RegenerationFailedException){throw;}
                            catch(Exception){ /* The affected metric reports the cached error during calculation. */ }
                        }
                var pending=ReviewInputs.Where(i=>i.Region!=null&&!i.Region.IsEmpty&&!reviewBindings.Any(b=>b.Key==i.Key)).ToList();
                try{CreateEditableInputs(reportProgress,pending);}
                catch(OperationCanceledException){throw;}
                catch(Autodesk.Revit.Exceptions.OperationCanceledException){throw;}
                catch(Autodesk.Revit.Exceptions.RegenerationFailedException){throw;}
                catch(Exception ex)
                {
                    foreach(var input in pending)input.Status="Область не построена: "+ex.Message;
                    current.Issue("AREA_INPUT_BUILD","Предупреждение",ex.Message);
                }
                foreach(var input in ReviewInputs.Where(i=>!reviewBindings.Any(b=>b.Key==i.Key)))areaBuildFailures[input.Key]=input.Status;
                // The next pass must read the actual persisted boundaries, never the pre-build numbers.
                measuredRoomRegions.Clear();roomNetRegions.Clear();roomCentreRegions.Clear();roomFloorRegions.Clear();roomFloorContours.Clear();
                shapes.Clear();DisposeSpatialCalculators();localSpatialVolumes.Clear();planarBodies?.Clear();floorWallFaces.Clear();floorFaceHeights.Clear();nativeWallLayers.Clear();
                clearanceGeometry.Clear();clearanceErrors.Clear();clearanceFaces.Clear();
                automaticAreaInputs=true;
                current.Issue("AREA_INPUT_WORKFLOW","Информация","Расчётные 2D-области создаются автоматически. Площади читаются из сохранённых границ этих областей; затем применяются исключения, коэффициенты и проверки. Разные нормативные обмеры имеют отдельные области. Непостроенные области не заменяются скрытым расчётом. Объёмы рассчитываются по 3D-геометрии.");
            }
        }
    }
}
