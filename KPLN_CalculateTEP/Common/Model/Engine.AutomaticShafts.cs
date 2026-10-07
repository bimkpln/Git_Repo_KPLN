using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public partial class Engine
        {
            // Project-independent naming convention. Do not replace this with names of particular buildings.
            internal const string AutomaticShaftNameToken="шахты";
            internal static bool IsAutomaticShaftName(string name)
            {return !string.IsNullOrEmpty(name)&&name.IndexOf(AutomaticShaftNameToken,StringComparison.OrdinalIgnoreCase)>=0;}

            internal static string ShaftContourIdentity(PlanarRegion region)
            {
                var rings=new List<string>();
                foreach(var path in region.Paths)
                {
                    var points=path.Select(p=>p.X.ToString(CultureInfo.InvariantCulture)+","+p.Y.ToString(CultureInfo.InvariantCulture)).ToList();
                    var variants=new List<string>();
                    for(int direction=0;direction<2;direction++)
                    {
                        if(direction==1)points.Reverse();
                        int first=Enumerable.Range(0,points.Count).OrderBy(i=>points[i],StringComparer.Ordinal).First();
                        variants.Add(string.Join(";",Enumerable.Range(0,points.Count).Select(i=>points[(first+i)%points.Count])));
                    }
                    rings.Add(variants.OrderBy(x=>x,StringComparer.Ordinal).First());
                }
                string text=string.Join("|",rings.OrderBy(x=>x,StringComparer.Ordinal));
                using(var hash=SHA256.Create())return BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(text))).Replace("-","");
            }

            internal static bool ShaftContoursConflict(PlanarRegion a,PlanarRegion b)
            {return ShaftContourIdentity(a)!=ShaftContourIdentity(b)&&RegionIntersection(a,b).Area*.09290304>.001;}

            internal static LevelSetting ResolveAutomaticShaftLevel(IEnumerable<LevelSetting> sourceLevels,string nativeKey,double elevation)
            {
                var matching=sourceLevels.Where(l=>l!=null&&Math.Abs(l.Elevation-elevation)<=FloorGeometryTolerance).ToList();
                if(matching.Count==0)
                    throw new InvalidOperationException("Границы шахт на отметке "+(elevation*.3048).ToString("0.######",CultureInfo.InvariantCulture)+" м не совпадают с расчётными уровнями источника.");
                // A group belongs to its native datum even when another building branch
                // has a datum at the same height. The geometry plane must still match.
                var native=string.IsNullOrEmpty(nativeKey)?new List<LevelSetting>():matching.Where(l=>l.Key==nativeKey).ToList();
                var candidates=native.Count>0?native:matching;
                if(candidates.Select(l=>l.Include+"|"+l.Kind+"|"+l.Above).Distinct(StringComparer.Ordinal).Count()!=1)
                {
                    string names=string.Join(", ",candidates.Select(l=>"«"+l.Name+"»"+
                        (string.IsNullOrWhiteSpace(l.Building)?"":" (корпус "+l.Building+")")+
                        (string.IsNullOrWhiteSpace(l.Section)?"":" (секция "+l.Section+")")).Distinct(StringComparer.Ordinal));
                    throw new InvalidOperationException((native.Count>0?"У родного уровня группы шахт есть противоречивые уточнения: ":
                        "На отметке границ шахт несколько уровней с разными настройками расчёта: ")+names+". Проверьте включение, вид этажа и наземность.");
                }
                return candidates.OrderBy(l=>l.Key,StringComparer.Ordinal).ThenBy(l=>l.Building,StringComparer.Ordinal)
                    .ThenBy(l=>l.Section,StringComparer.Ordinal).ThenBy(l=>l.Name,StringComparer.Ordinal).First();
            }

            private List<Curve> ShaftGroupCurves(Group group,Source source,HashSet<ElementId> visited)
            {
                var curves=new List<Curve>();
                foreach(var id in group.GetMemberIds())
                {
                    if(!visited.Add(id))continue;
                    var member=source.Document.GetElement(id);
                    if(member is Group)curves.AddRange(ShaftGroupCurves((Group)member,source,visited));
                    else if(member is CurveElement)
                    {
                        if(!PhaseAccepted(source,member))continue;
                        curves.Add(((CurveElement)member).GeometryCurve.CreateTransformed(source.Transform));
                    }
                    else throw new InvalidOperationException("В группе есть объект «"+member?.Name+"» (ID "+IDHelper.ElIdValue(id)+"), который не является линией границы. Габарит группы не подставлен вместо шахт.");
                }
                return curves;
            }

            private void CollectAutomaticShafts(List<Record> records)
            {
                if(!Config.Metrics.Any(m=>m.Enabled&&(RoomEnvelopeTotal((Indicator)Enum.Parse(typeof(Indicator),m.Key))||RoomEnvelopePart((Indicator)Enum.Parse(typeof(Indicator),m.Key))||RoomAreaMetric((Indicator)Enum.Parse(typeof(Indicator),m.Key)))))return;
                var added=new List<Record>();int duplicates=0,groups=0;
                var seen=new HashSet<string>(StringComparer.Ordinal);
                foreach(var source in Sources.Where(s=>s.Loaded&&s.LoadError==null&&s.Mode!="exclude"&&Math.Abs(s.Transform.BasisZ.DotProduct(XYZ.BasisZ)-1)<1e-8))
                    foreach(var group in source.Elements.OfType<Group>().Where(g=>IsAutomaticShaftName(g.Name)).OrderBy(g=>IDHelper.ElIdValue(g.Id)))
                    {
                        if(Config.Grouping=="selection"&&!selection.Contains(IDHelper.ElIdValue(source.RootLink??group.Id)))continue;
                        if(group.GroupId!=ElementId.InvalidElementId)
                        {
                            // The outer named group already contributes its nested boundaries.
                            var parent=source.Document.GetElement(group.GroupId) as Group;bool nested=false;
                            while(parent!=null){if(IsAutomaticShaftName(parent.Name)){nested=true;break;}parent=source.Document.GetElement(parent.GroupId) as Group;}
                            if(nested)continue;
                        }
                        var context=new Record{Source=source,Element=group};
                        try
                        {
                            if(!PhaseAccepted(source,group))continue;
                            Progress("Шахты по имени группы: "+source.Name+" / "+group.Name);groups++;
                            var curves=ShaftGroupCurves(group,source,new HashSet<ElementId>());
                            if(curves.Count==0)throw new InvalidOperationException("Нет линий замкнутой границы шахт.");
                            double z=curves[0].GetEndPoint(0).Z;
                            var geometry=NativeBoundaryRegion(curves,z);
                            if(geometry.IsEmpty)throw new InvalidOperationException("Граница шахт не образует положительную площадь.");
                            var nativeLevel=ModelLevel(group);
                            var level=ResolveAutomaticShaftLevel(Config.Levels.Where(l=>l.Key.StartsWith(source.Key+"/",StringComparison.Ordinal)&&l.Key.Substring(source.Key.Length+1).IndexOf('/')<0),
                                nativeLevel==null?null:source.Key+"/"+nativeLevel.UniqueId,z);context.Level=level;
                            if(!level.Include||level.Kind=="exclude")continue;
                            var peers=records.Where(r=>r.Element is Room&&r.Source==source&&r.Level!=null&&Math.Abs(r.Z-z)<=FloorGeometryTolerance).ToList();
                            if(peers.Count==0)peers=records.Where(r=>r.Element is Room&&r.Source==source).ToList();
                            if(peers.Count==0)peers=records.Where(r=>r.Element is Room&&r.Source.Mode==source.Mode).ToList();
                            string assignedBuilding;
                            try{assignedBuilding=AssignedBuilding(source,group);}
                            catch(Exception ex){Notice("AUTO_SHAFT_BUILDING","Ошибка",ex.Message,context);continue;}
                            if(!string.IsNullOrWhiteSpace(assignedBuilding))
                            {
                                peers=peers.Where(r=>Eq(r.Building,assignedBuilding)).ToList();
                                if(peers.Count==0)
                                {
                                    string profile=Config.Profile=="by-source"?source.Profile:Config.Profile;
                                    if(string.IsNullOrEmpty(profile)||profile.StartsWith("by-")||profile=="high-mixed")
                                    {Notice("AUTO_SHAFT_BUILDING","Ошибка","Корпус задан соответствием, но не определён профиль здания. Укажите профиль в показателях.",context);continue;}
                                    peers.Add(new Record{Source=source,Element=group,Level=level,Building=assignedBuilding,Profile=profile,BuildingClass=profile.Contains("public")?"nonresidential":"residential",Part="auto"});
                                }
                            }
                            if(peers.Count==0||peers.Any(r=>string.IsNullOrWhiteSpace(r.Building))||peers.Select(r=>r.Building+"|"+r.Profile+"|"+r.BuildingClass).Distinct().Count()!=1)
                            {
                                string reason=peers.Count==0?"Нет принятых в расчёт помещений, по которым можно определить корпус шахт. Сначала устраните причины исключения помещений (в том числе ПОМ_Корпус) либо подтвердите единый корпус, если это верно для модели.":
                                    "У помещений источника отсутствует корпус либо указаны разные корпуса / типы зданий. Уточните распределение помещений по корпусам; назначать шахту произвольному корпусу нельзя.";
                                Notice("AUTO_SHAFT_BUILDING","Ошибка",reason+" Границы группы прочитаны; это ошибка определения корпуса, а не построения контура.",context);
                                continue;
                            }
                            foreach(var paths in geometry.Components())
                            {
                                var region=new PlanarRegion(paths);string identity=ShaftContourIdentity(region);
                                var record=peers[0].Copy();record.Source=source;record.Element=group;record.Level=level;
                                record.Role="shaft";record.Apartment="";record.Override=null;record.Manual=false;record.Factor=null;
                                record.Part=peers.Select(r=>r.Part).Distinct().Count()==1?peers[0].Part:"auto";
                                record.Section=peers.Select(r=>r.Section??"").Distinct().Count()==1?peers[0].Section:"";
                                record.AutomaticShaftRegion=region;record.AutomaticShaftPart=identity;
                                // The global contour establishes vertical continuity; names identify candidates only.
                                record.Vertical="auto-shaft/"+record.Building+"/"+identity;
                                string key=source.Mode+"/"+record.Building+"/"+Math.Round(z,6).ToString("R",CultureInfo.InvariantCulture)+"/"+identity;
                                if(!seen.Add(key)){duplicates++;continue;}
                                added.Add(record);
                            }
                        }
                        catch(OperationCanceledException){throw;}
                        catch(Exception ex){Notice("AUTO_SHAFT_GEOMETRY","Ошибка",ex.Message,context);}
                    }
                // Section metadata may differ by floor. Exact global contours in the same building still identify one shaft.
                foreach(var shaft in added.GroupBy(r=>r.Vertical))
                    if(shaft.Select(r=>r.Section??"").Distinct().Count()>1)foreach(var record in shaft)record.Section="";
                foreach(var scope in added.GroupBy(r=>r.Source.Mode+"/"+r.Building))
                {
                    var stacks=scope.GroupBy(r=>r.Vertical).Select(g=>g.ToList()).ToList();
                    var neighbours=Enumerable.Range(0,stacks.Count).Select(i=>new HashSet<int>()).ToList();
                    for(int i=0;i<stacks.Count;i++)for(int j=i+1;j<stacks.Count;j++)
                        if(ShaftContoursConflict(stacks[i][0].AutomaticShaftRegion,stacks[j][0].AutomaticShaftRegion))
                        {
                            double overlapArea=RegionIntersection(stacks[i][0].AutomaticShaftRegion,stacks[j][0].AutomaticShaftRegion).Area*.09290304;
                            neighbours[i].Add(j);neighbours[j].Add(i);
                            foreach(var record in stacks[i].Concat(stacks[j]).Where(r=>stacks[i].Any(a=>Math.Abs(a.Z-r.Z)<=FloorGeometryTolerance)&&stacks[j].Any(b=>Math.Abs(b.Z-r.Z)<=FloorGeometryTolerance)))
                            {
                                record.AutomaticShaftError="На одной расчётной отметке "+(record.Z*.3048).ToString("+0.000;-0.000;0.000")+
                                    " м пересекаются разные контуры шахт. Площадь пересечения "+overlapArea.ToString("0.###")+
                                    " м². Нельзя отличить смещённую копию группы от другого расчётного контура; проверьте наложенные группы и их границы. Остальные отметки рассчитываются отдельно.";
                                var conflicting=stacks[i].Concat(stacks[j]).Where(r=>Math.Abs(r.Z-record.Z)<=FloorGeometryTolerance).Select(r=>IDHelper.ElIdValue(r.Element.Id).ToString()).Distinct().OrderBy(id=>id);
                                Notice("AUTO_SHAFT_CONTINUITY","Ошибка",record.AutomaticShaftError+" Этаж: «"+record.Level.Name+"». Группы ID: "+string.Join(", ",conflicting)+".",record);
                            }
                        }
                    var visited=new HashSet<int>();
                    for(int i=0;i<stacks.Count;i++)
                    {
                        if(visited.Contains(i))continue;var queue=new Queue<int>();var component=new List<int>();queue.Enqueue(i);
                        while(queue.Count>0){int index=queue.Dequeue();if(!visited.Add(index))continue;component.Add(index);foreach(int next in neighbours[index])queue.Enqueue(next);}
                        if(component.Count<2)continue;
                        var members=component.SelectMany(index=>stacks[index]).ToList();
                        // A changing footprint is one shaft only if every floor has one unambiguous contour.
                        if(members.GroupBy(r=>Math.Round(r.Z,6)).Any(g=>g.Select(r=>r.AutomaticShaftPart).Distinct().Count()>1))continue;
                        string identity=members.Select(r=>r.Vertical).OrderBy(x=>x,StringComparer.Ordinal).First();
                        foreach(var record in members){record.Vertical=identity;record.Section="";}
                    }
                }
                records.AddRange(added);
                if(groups>0)Notice("AUTO_SHAFTS","Информация","Группы с «"+AutomaticShaftNameToken+"» в имени: "+groups+". Отдельных поэтажных контуров: "+added.Count+"; совпадающих повторов исключено: "+duplicates+". ID вертикальных шахт определены по совпадающим контурам в плане; параметры модели не записывались.");
            }

            private PlanarRegion AutomaticShaftPlan(Record record)
            {
                if(record.AutomaticShaftError!=null)throw new InvalidOperationException(record.AutomaticShaftError);
                return ReviewRegion("shaft/"+record.Key,record,"Шахта - границы группы",record.Z,()=>record.AutomaticShaftRegion);
            }

            private PlanarRegion AutomaticShaftFloorRegion(double elevation)
            {
                return automaticGeometryRecords.Where(r=>r.AutomaticShaftRegion!=null&&r.Source.Mode=="include"&&r.Level!=null&&r.Level.Include&&Math.Abs(r.Z-elevation)<=FloorGeometryTolerance)
                    .Aggregate(PlanarRegion.Empty,(sum,r)=>RegionUnion(sum,AutomaticShaftPlan(r)));
            }
        }
    }
}
