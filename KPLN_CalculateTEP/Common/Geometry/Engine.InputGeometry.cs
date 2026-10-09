using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public partial class Engine
        {
            private static double[] ElementBounds(Element element, Transform transform)
            {
                var box = element?.get_BoundingBox(null);
                if (box == null) return null;
                var placement = transform.Multiply(box.Transform);
                var points = Enumerable.Range(0, 8).Select(i => placement.OfPoint(new XYZ(
                    (i & 1) == 0 ? box.Min.X : box.Max.X,
                    (i & 2) == 0 ? box.Min.Y : box.Max.Y,
                    (i & 4) == 0 ? box.Min.Z : box.Max.Z))).ToList();
                return new[] { points.Min(p => p.X), points.Min(p => p.Y), points.Min(p => p.Z),
                    points.Max(p => p.X), points.Max(p => p.Y), points.Max(p => p.Z) };
            }

            private static string DescribeGeometryElement(Element element, Transform transform, string source)
            {
                if (element == null) return "Источник: " + source + "; исходный элемент недоступен";
                var type = element.Document.GetElement(element.GetTypeId());
                string text = "Источник: " + source + "; ID " + IDHelper.ElIdValue(element.Id) +
                    "; категория: " + (element.Category?.Name ?? "без категории") +
                    "; класс: " + element.GetType().Name + "; тип: " + (type?.Name ?? element.Name ?? "не задан");
                var bounds = ElementBounds(element, transform);
                if (bounds != null) text += "; габарит XYZ: " + string.Join(" x ", Enumerable.Range(0, 3)
                    .Select(i => ((bounds[i + 3] - bounds[i]) * .3048).ToString("0.###", CultureInfo.InvariantCulture))) +
                    " м; Z: " + (bounds[2] * .3048).ToString("0.###", CultureInfo.InvariantCulture) +
                    " ... " + (bounds[5] * .3048).ToString("0.###", CultureInfo.InvariantCulture) + " м";
                return text;
            }

            private static PlanarRegion BoundsRegion(double[] bounds)
            {
                if (bounds == null || bounds[3] <= bounds[0] || bounds[4] <= bounds[1]) return PlanarRegion.Empty;
                return PlanarRegion.FromRings(new[] { new[] {
                    new[] { bounds[0], bounds[1] }, new[] { bounds[3], bounds[1] },
                    new[] { bounds[3], bounds[4] }, new[] { bounds[0], bounds[4] } } });
            }

            // This is an ownership guard, NOT the footprint and never a clipping polygon.
            // A large site slab cannot become part of a building simply because the source has one building.
            private PlanarRegion FootprintSourceScope(IEnumerable<Record> records, string building)
            {
                var bounds = new List<double[]>(); var visited = new HashSet<string>();
                foreach (var record in records.Where(r => r.Element is Room && Eq(r.Building, building) && r.Source.Mode != "exclude" && r.Level != null && r.Level.Include))
                {
                    var roomBounds = ElementBounds(record.Element, record.Source.Transform);
                    if (roomBounds != null) bounds.Add(roomBounds);
                    var rings = ((Room)record.Element).GetBoundarySegments(new SpatialElementBoundaryOptions { SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish });
                    if (rings == null) continue;
                    foreach (var segment in rings.SelectMany(r => r))
                    {
                        string key = record.Source.Key + "/" + (segment.ElementId==null?-1:IDHelper.ElIdValue(segment.ElementId)) + "/" + (segment.LinkElementId==null?-1:IDHelper.ElIdValue(segment.LinkElementId));
                        if (!visited.Add(key)) continue;
                        var element = record.Element.Document.GetElement(segment.ElementId);
                        var transform = record.Source.Transform;
                        var link = element as RevitLinkInstance;
                        if (link != null) { element = link.GetLinkDocument()?.GetElement(segment.LinkElementId); transform = transform.Multiply(link.GetTotalTransform()); }
                        var wall = element as Wall ?? (element as FamilyInstance)?.Host as Wall;
                        if (wall == null) continue;
                        var wallBounds = ElementBounds(wall, transform);
                        if (wallBounds != null) bounds.Add(wallBounds);
                    }
                }
                if (bounds.Count == 0) return PlanarRegion.Empty;
                return BoundsRegion(new[] { bounds.Min(b => b[0]), bounds.Min(b => b[1]), bounds.Min(b => b[2]),
                    bounds.Max(b => b[3]), bounds.Max(b => b[4]), bounds.Max(b => b[5]) });
            }

            private bool AcceptFootprintRegion(Record record, Solid source, PlanarRegion region, PlanarRegion scope)
            {
                var box = BoundsRegion(Bounds(source));
                string problem=FootprintExtentProblem(region,box,scope,record.Role=="structure"&&record.Override!=true);
                string description = DescribeGeometryElement(record.Element, record.Source.Transform, record.Source.Name) +
                    "; площадь проекции: " + (region.Area * .09290304).ToString("0.###", CultureInfo.InvariantCulture) + " м²";
                if (problem=="FOOTPRINT_PROJECTION_BOUNDS")
                {
                    Notice("FOOTPRINT_PROJECTION_BOUNDS", "Ошибка", "Проекция вышла за габариты исходного тела. Объект не включён; требуется проверка вычисления проекции. " + description,
                        record, Indicator.Footprint.ToString(), "Не изменяйте модель по этой ошибке: нужна проверка геометрического алгоритма по указанному объекту.");
                    return false;
                }
                // Explicit calculation envelopes retain their deliberately assigned extent.
                if (problem=="FOOTPRINT_SOURCE_SCOPE")
                {
                    double outsideScope = RegionDifference(region, scope).Area;
                    Notice("FOOTPRINT_SOURCE_SCOPE", "Ошибка", "Не подтверждена принадлежность всей конструкции зданию: проекция выходит за габарит размещённых помещений и ограничивающих их стен на " +
                            (outsideScope * .09290304).ToString("0.###", CultureInfo.InvariantCulture) + " м². Объект целиком пропущен, площадь не обрезана автоматически. " + description,
                            record, Indicator.Footprint.ToString(), "Проверьте назначение объекта: это может быть элемент участка или выступающая часть здания. До подтверждения результат застройки неполный; параметр корпуса у конструкции не требуется.");
                    return false;
                }
                return true;
            }

            private static string FootprintExtentProblem(PlanarRegion region,PlanarRegion box,PlanarRegion scope,bool requireScope)
            {
                double tolerance=Math.Max(1e-7,(box.Perimeter+region.Perimeter)*(ArcChordTolerance*2+2/PlanarRegion.Scale));
                if(RegionDifference(region,box).Area>tolerance||region.Area>box.Area+tolerance)return "FOOTPRINT_PROJECTION_BOUNDS";
                if(requireScope&&(scope.IsEmpty||RegionDifference(region,scope).Area>tolerance))return "FOOTPRINT_SOURCE_SCOPE";
                return null;
            }

            public static bool IsCalculationFrameEdge(double ax, double ay, double bx, double by,
                double left, double bottom, double right, double top, double tolerance)
            {
                if (!(right > left) || !(top > bottom) || tolerance < 0 || ax==bx && ay==by) return false;
                bool inside = Math.Min(ax,bx)>=left-tolerance && Math.Max(ax,bx)<=right+tolerance &&
                    Math.Min(ay,by)>=bottom-tolerance && Math.Max(ay,by)<=top+tolerance;
                return inside && ((Math.Abs(ax-left)<=tolerance && Math.Abs(bx-left)<=tolerance) ||
                    (Math.Abs(ax-right)<=tolerance && Math.Abs(bx-right)<=tolerance) ||
                    (Math.Abs(ay-bottom)<=tolerance && Math.Abs(by-bottom)<=tolerance) ||
                    (Math.Abs(ay-top)<=tolerance && Math.Abs(by-top)<=tolerance));
            }

            private sealed class ResolvedShellBoundary
            {
                internal readonly List<Record> Supports=new List<Record>();
                internal PlanarRegion Section;
            }
            private readonly Dictionary<string,PlanarRegion> boundarySupportRegions=new Dictionary<string,PlanarRegion>();
            private string ShellBoundaryKey(Record record)
            { return (record.Source.Document.Equals(doc)?"host":IDHelper.ElIdValue(record.Source.RootLink).ToString())+"/"+IDHelper.ElIdValue(record.Element.Id); }

            private static bool IsPhysicalShellCandidate(Element element)
            {
                if(element is HostObject)return true;
                var instance=element as FamilyInstance;
                if(instance==null||instance.Category==null)return false;
                long category=IDHelper.ElIdValue(instance.Category.Id);
                if(category==(long)BuiltInCategory.OST_StructuralColumns||category==(long)BuiltInCategory.OST_Columns||
                    category==(long)BuiltInCategory.OST_CurtainWallPanels||category==(long)BuiltInCategory.OST_CurtainWallMullions)return true;
                var parameter=instance.get_Parameter(BuiltInParameter.WALL_ATTR_ROOM_BOUNDING);
                // An explicit instance value, including false, takes precedence over its type.
                if(parameter==null)parameter=instance.Symbol?.get_Parameter(BuiltInParameter.WALL_ATTR_ROOM_BOUNDING);
                return parameter!=null&&parameter.HasValue&&parameter.StorageType==StorageType.Integer&&parameter.AsInteger()==1;
            }

            private List<Record> PhysicalShellCandidates(List<Record> walls,Record floor,double elevation)
            {
                var result=new List<Record>(walls);var known=new HashSet<string>(walls.Select(r=>r.Key));
                foreach(var source in Sources.Where(s=>s.Loaded&&s.LoadError==null&&s.Mode=="include"))
                    foreach(var element in source.Elements.Where(e=>!(e is Wall)&&IsPhysicalShellCandidate(e)&&PhaseAccepted(source,e)&&CrossesFloor(e,source.Transform,elevation)))
                    {
                        var record=new Record{Source=source,Element=element,Level=floor.Level,Building=floor.Building,Role="structure"};
                        if(known.Add(record.Key))result.Add(record);
                    }
                return result;
            }

            private PlanarRegion PhysicalShellSection(Record record,double elevation,Indicator indicator)
            {
                string key=record.Key+"/"+elevation.ToString("R",CultureInfo.InvariantCulture);
                PlanarRegion region;
                if(!boundarySupportRegions.TryGetValue(key,out region))
                {
                    region=PlanarRegion.Empty;
                    foreach(var solid in VolumeSolids(record,indicator.ToString()))region=RegionUnion(region,SectionRegion(solid,elevation));
                    boundarySupportRegions[key]=region;
                }
                return region;
            }

            private PlanarRegion FloorMaterialRegion(List<Record> candidates,double elevation,Indicator indicator,PlanarRegion scope)
            {
                var result=PlanarRegion.Empty;
                foreach(var record in candidates)
                {
                    // Horizontal slabs must never fill a courtyard or an unclassified room.
                    if(record.Element is Floor||record.Element is RoofBase||record.Element is Ceiling)continue;
                    var bounds=ElementBounds(record.Element,record.Source.Transform);
                    if(bounds==null||RegionIntersection(scope,BoundsRegion(bounds)).IsEmpty)continue;
                    Progress("2D-сечения ограждений: ID "+IDHelper.ElIdValue(record.Element.Id));
                    try{result=RegionUnion(result,RegionIntersection(scope,PhysicalShellSection(record,elevation,indicator)));}
                    catch(OperationCanceledException){throw;}
                    catch(Autodesk.Revit.Exceptions.OperationCanceledException){throw;}
                    catch(Autodesk.Revit.Exceptions.RegenerationFailedException){throw;}
                    catch(Exception ex){throw new UnconfirmedEnvelopeException("Не подтверждено сечение ограждения ID "+IDHelper.ElIdValue(record.Element.Id)+": "+ex.Message);}
                }
                return result;
            }

            // Native exterior boundaries already contain the selected finish/core geometry.
            // Resolve identity only for source/phase ownership and courtyard checks; do not
            // force every segment back through a physical-solid reconstruction.
            private ResolvedShellBoundary ReadNativeShellBoundary(BoundarySegment segment,List<Record> candidates,Record floor,Indicator indicator)
            {
                var result=new ResolvedShellBoundary();var element=doc.GetElement(segment.ElementId);
                string root="host";var link=element as RevitLinkInstance;
                if(link!=null)
                {
                    root=IDHelper.ElIdValue(link.Id).ToString();
                    if(!Sources.Any(s=>s.Loaded&&s.LoadError==null&&s.Mode=="include"&&s.RootLink==link.Id))
                        throw new UnconfirmedEnvelopeException("Граница принадлежит невыбранной связи: "+link.Name);
                    element=link.GetLinkDocument()?.GetElement(segment.LinkElementId);
                }
                if(element==null)
                {
                    var curve=segment.GetCurve();var a=curve.GetEndPoint(0);var b=curve.GetEndPoint(1);
                    Notice("NATIVE_BOUNDARY_UNRESOLVED","Предупреждение","Использован участок замкнутой границы Revit без доступного ID элемента. Режим границы и отметка заданы расчётом. Проверьте происхождение участка XY: "+
                        (a.X*.3048).ToString("0.###",CultureInfo.InvariantCulture)+", "+(a.Y*.3048).ToString("0.###",CultureInfo.InvariantCulture)+" -> "+
                        (b.X*.3048).ToString("0.###",CultureInfo.InvariantCulture)+", "+(b.Y*.3048).ToString("0.###",CultureInfo.InvariantCulture)+" м.",floor,indicator.ToString());
                    return result;
                }
                var source=Sources.FirstOrDefault(s=>s.Loaded&&s.LoadError==null&&s.Mode=="include"&&s.Document.Equals(element.Document)&&
                    (s.Document.Equals(doc)?"host":IDHelper.ElIdValue(s.RootLink).ToString())==root);
                if(source==null||!PhaseAccepted(source,element))throw new UnconfirmedEnvelopeException("Граница принадлежит невыбранному источнику / стадии, ID "+IDHelper.ElIdValue(element.Id));
                if(element is ModelCurve)throw new UnconfirmedEnvelopeException("Наружная оболочка ограничена разделителем помещений ID "+IDHelper.ElIdValue(element.Id)+". Его нормативное положение не подтверждено конструкцией.");
                var wall=element as Wall??(element as FamilyInstance)?.Host as Wall;
                if(wall!=null)
                {
                    var members=new HashSet<string>(WallMembers(wall).Select(w=>w.UniqueId));
                    result.Supports.AddRange(candidates.Where(r=>r.Source.Key==source.Key&&r.Element is Wall&&members.Contains(r.Element.UniqueId)));
                    if(result.Supports.Count==0)throw new UnconfirmedEnvelopeException("Стена границы отсутствует на расчётной отметке / стадии, ID "+IDHelper.ElIdValue(wall.Id));
                }
                else result.Supports.Add(new Record{Source=source,Element=element,Level=floor.Level,Building=floor.Building,Role="structure"});
                return result;
            }

            private ResolvedShellBoundary ResolveShellBoundary(BoundarySegment segment,List<Record> candidates,double elevation,Record floor,Indicator indicator)
            {
                var element=segment.ElementId==null?null:doc.GetElement(segment.ElementId);string sourceName=doc.Title,root="host";
                var transform=Transform.Identity;
                var link=element as RevitLinkInstance;
                if(link!=null)
                {
                    sourceName=link.Name;root=IDHelper.ElIdValue(link.Id).ToString();transform=link.GetTotalTransform();
                    element=link.GetLinkDocument()?.GetElement(segment.LinkElementId);
                }
                string identity="ID границы "+(segment.ElementId==null?"нет":IDHelper.ElIdValue(segment.ElementId).ToString())+
                    "; ID в связи "+(segment.LinkElementId==null?"нет":IDHelper.ElIdValue(segment.LinkElementId).ToString());
                var description=identity+"; "+DescribeGeometryElement(element,transform,sourceName);
                var host=element as Wall??(element as FamilyInstance)?.Host as Wall;
                if(host!=null)
                {
                    var members=WallMembers(host).Select(w=>w.UniqueId).ToList();
                    var matches=candidates.Where(w=>w.Element is Wall&&members.Contains(w.Element.UniqueId)&&
                        (w.Source.Document.Equals(doc)?"host":IDHelper.ElIdValue(w.Source.RootLink).ToString())==root).ToList();
                    if(matches.Count==1)
                    {
                        if(!(element is Wall))Notice("BOUNDARY_HOST_WALL","Информация","Граница семейства сопоставлена с его стеной-основой ID "+IDHelper.ElIdValue(matches[0].Element.Id)+". "+description,floor,indicator.ToString());
                        var result=new ResolvedShellBoundary();result.Supports.Add(matches[0]);return result;
                    }
                    if(matches.Count==0)throw new UnconfirmedEnvelopeException("Стена границы не входит в выбранные источники / стадию на отметке пола. "+description);
                }
                // An actual model element returned by the Room API is admissible even when its
                // category was not preselected, but annotation/separation lines are not solids.
                if(element!=null&&!(element is ModelCurve)&&!(element is SpatialElement)&&element.Category?.CategoryType==CategoryType.Model)
                {
                    var source=Sources.FirstOrDefault(s=>s.Loaded&&s.Mode=="include"&&s.LoadError==null&&s.Document.Equals(element.Document)&&
                        (s.Document.Equals(doc)?"host":IDHelper.ElIdValue(s.RootLink).ToString())==root);
                    if(source==null||!PhaseAccepted(source,element))throw new UnconfirmedEnvelopeException("Физическая граница принадлежит невыбранному источнику / стадии. "+description);
                    if(!candidates.Any(r=>r.Source.Key==source.Key&&r.Element.UniqueId==element.UniqueId))
                        candidates.Add(new Record{Source=source,Element=element,Level=floor.Level,Building=floor.Building,Role="structure"});
                }
                var points=CurvePoints(segment.GetCurve());
                description+="; начало XY: "+(points[0][0]*.3048).ToString("0.###",CultureInfo.InvariantCulture)+", "+(points[0][1]*.3048).ToString("0.###",CultureInfo.InvariantCulture)+
                    " м; конец XY: "+(points.Last()[0]*.3048).ToString("0.###",CultureInfo.InvariantCulture)+", "+(points.Last()[1]*.3048).ToString("0.###",CultureInfo.InvariantCulture)+" м";
                double minX=points.Min(p=>p[0]),minY=points.Min(p=>p[1]),maxX=points.Max(p=>p[0]),maxY=points.Max(p=>p[1]);
                var supported=new List<Tuple<Record,PlanarRegion>>();var failures=new List<string>();
                Func<Record,bool> isActual=r=>element!=null&&r.Element.UniqueId==element.UniqueId&&r.Element.Document.Equals(element.Document)&&
                    (r.Source.Document.Equals(doc)?"host":IDHelper.ElIdValue(r.Source.RootLink).ToString())==root;
                foreach(var candidate in candidates.OrderByDescending(isActual))
                {
                    var bounds=ElementBounds(candidate.Element,candidate.Source.Transform);
                    // An anonymous junction can be covered jointly by two adjacent elements.
                    if(bounds==null||maxX<bounds[0]-FloorGeometryTolerance||maxY<bounds[1]-FloorGeometryTolerance||minX>bounds[3]+FloorGeometryTolerance||minY>bounds[4]+FloorGeometryTolerance)continue;
                    try
                    {
                        var region=PhysicalShellSection(candidate,elevation,indicator);
                        if(!region.IsEmpty)supported.Add(Tuple.Create(candidate,region));
                        if(isActual(candidate)&&
                            PlanarRegion.SupportsBoundaryPolyline(candidate.Element is Wall?region:PlanarRegion.Empty,candidate.Element is Wall?new PlanarRegion[0]:new[]{region},points,ArcChordTolerance*2,planarCheckpoint))
                        {supported.Clear();supported.Add(Tuple.Create(candidate,region));break;}
                    }
                    catch(OperationCanceledException){throw;}
                    catch(Autodesk.Revit.Exceptions.OperationCanceledException){throw;}
                    catch(Autodesk.Revit.Exceptions.RegenerationFailedException){throw;}
                    catch(Exception ex){failures.Add("ID "+IDHelper.ElIdValue(candidate.Element.Id)+": "+ex.Message);}
                }
                Func<IEnumerable<Tuple<Record,PlanarRegion>>,PlanarRegion> union=items=>items.Aggregate(PlanarRegion.Empty,(r,item)=>RegionUnion(r,item.Item2));
                Func<IEnumerable<Tuple<Record,PlanarRegion>>,bool> covers=items=>PlanarRegion.SupportsBoundaryPolyline(
                    union(items.Where(i=>i.Item1.Element is Wall)),items.Where(i=>!(i.Item1.Element is Wall)).Select(i=>i.Item2),points,ArcChordTolerance*2,planarCheckpoint);
                if(!covers(supported))
                    throw new UnconfirmedEnvelopeException("Вся граница не подтверждена сечениями физических ограждений на отметке пола. Наличие ID -1 само по себе не считается разрывом; участок проверен по объединению соседних элементов. "+description+
                        "; проверены элементы: "+string.Join(", ",supported.Select(w=>IDHelper.ElIdValue(w.Item1.Element.Id)))+
                        (failures.Count>0?"; ошибки сечений: "+string.Join(" | ",failures.Take(3)):""));
                // Keep only necessary contributors. Bounding-box neighbours must not be marked
                // seen: that could hide another external wall or an unexamined courtyard.
                for(int i=supported.Count-1;i>=0&&supported.Count>1;i--)
                    if(covers(supported.Where((r,index)=>index!=i)))supported.RemoveAt(i);
                var resolved=new ResolvedShellBoundary{Section=union(supported)};
                resolved.Supports.AddRange(supported.Select(s=>s.Item1));
                Notice("BOUNDARY_PHYSICAL","Информация","Граница подтверждена фактическими сечениями ограждений: "+
                    string.Join("; ",resolved.Supports.Select(r=>DescribeGeometryElement(r.Element,r.Source.Transform,r.Source.Name)))+". "+identity,floor,indicator.ToString());
                return resolved;
            }

            private double PhysicalBoundaryInset(ResolvedShellBoundary boundary,Curve curve,double elevation,Indicator indicator)
            {
                if(boundary.Supports.Count==1&&boundary.Supports[0].Element is Wall)
                {
                    var record=boundary.Supports[0];var wall=(Wall)record.Element;
                    var a=WallFaceAtFloor(wall,record.Source.Transform,curve,elevation,ShellLayerType.Exterior);
                    var b=WallFaceAtFloor(wall,record.Source.Transform,curve,elevation,ShellLayerType.Interior);
                    if(Math.Min(Math.Abs(a.Offset),Math.Abs(b.Offset))>FloorGeometryTolerance*10)
                        throw new InvalidOperationException("Не подтверждена сторона оболочки у стены ID "+IDHelper.ElIdValue(wall.Id));
                    return Math.Abs(a.Offset)<Math.Abs(b.Offset)?b.Offset:a.Offset;
                }
                if(!(curve is Line))throw new UnconfirmedEnvelopeException("Наружная граница физического элемента подтверждена, но постоянное смещение его криволинейной внутренней стороны не установлено. ID: "+string.Join(", ",boundary.Supports.Select(r=>IDHelper.ElIdValue(r.Element.Id))));
                var section=boundary.Section??boundary.Supports.Aggregate(PlanarRegion.Empty,(r,item)=>RegionUnion(r,PhysicalShellSection(item,elevation,indicator)));
                var p=curve.GetEndPoint(0);var q=curve.GetEndPoint(1);
                try{return section.UniformBoundaryInset(new[]{p.X,p.Y},new[]{q.X,q.Y});}
                catch(InvalidOperationException ex){throw new UnconfirmedEnvelopeException(ex.Message+" ID ограждений: "+string.Join(", ",boundary.Supports.Select(r=>IDHelper.ElIdValue(r.Element.Id))));}
            }
        }
    }
}
