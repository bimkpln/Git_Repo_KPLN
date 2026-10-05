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
            private readonly Dictionary<string, bool> roomHeightSections = new Dictionary<string, bool>();
            private static bool RoomAreaMetric(Indicator metric)
            {
                return OneOf(metric.ToString(), "ApartmentsTotal", "ApartmentsHeated", "PublicRooms", "PublicCalculated");
            }

            private Solid RoomAreaBoundary(Record r)
            {
                string key = r.Key + "/native-area-boundary";
                Solid result;
                if (shapes.TryGetValue(key, out result)) return result;
                RequireHorizontalTransform(r.Source.Transform);
                var boundaries = ((Room)r.Element).GetBoundarySegments(new SpatialElementBoundaryOptions { SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish });
                if (boundaries == null || boundaries.Count == 0) throw new InvalidOperationException("У помещения отсутствует замкнутая чистовая граница.");
                var loops = boundaries.Select(ring => FlatLoop(ring.Select(s => s.GetCurve().CreateTransformed(r.Source.Transform)))).ToList();
                result = PlanFromSectionLoops(loops);
                if (result == null) throw new InvalidOperationException("Пустой контур помещения.");
                shapes[key] = result;
                return result;
            }

            private bool RoomRequiresHeightSection(Record r)
            {
                if (r.Role == "under-stair" || Number(r.Element, "slope").HasValue) return true;
                bool cached;
                if (roomHeightSections.TryGetValue(r.Key, out cached)) return cached;
                bool variable=RoomHasVariableWalls(r);
                // Curtain walls, stacked walls and non-wall boundaries need not change the room section.
                // Use the actual room solid as proof; do not infer a varying section from category alone.
                if(variable&&RoomIsVerticalPrism(r))variable=false;
                return roomHeightSections[r.Key] = variable;
            }
            private bool RoomIsVerticalPrism(Record r)
            {
                try
                {
                    var solid=SpatialVolume(r);var box=solid.GetBoundingBox();
                    var transform=box.Transform;
                    var heights=Enumerable.Range(0,8).Select(i=>transform.OfPoint(new XYZ((i&1)==0?box.Min.X:box.Max.X,(i&2)==0?box.Min.Y:box.Max.Y,(i&4)==0?box.Min.Z:box.Max.Z)).Z).ToList();
                    double bottom=heights.Min(),top=heights.Max();
                    double floor=RoomFloorElevation(r);
                    double native=((Room)r.Element).Level.ProjectElevation+((Room)r.Element).Level.get_Parameter(BuiltInParameter.LEVEL_ROOM_COMPUTATION_HEIGHT).AsDouble();
                    native=r.Source.Transform.OfPoint(new XYZ(0,0,native)).Z;
                    if(top-bottom<=FloorGeometryTolerance||floor<bottom-FloorGeometryTolerance||floor>=top||native<bottom||native>top)return false;
                    foreach(Face face in solid.Faces)
                    {
                        var plane=face as PlanarFace;
                        if(plane!=null)
                        {
                            double z=Math.Abs(plane.FaceNormal.Z);
                            if(z<1e-10)continue;
                            if(z>1-1e-10&&(Math.Abs(plane.Origin.Z-bottom)<FloorGeometryTolerance||Math.Abs(plane.Origin.Z-top)<FloorGeometryTolerance))continue;
                            return false;
                        }
                        var cylinder=face as CylindricalFace;
                        if(cylinder!=null&&Math.Abs(Math.Abs(cylinder.Axis.Z)-1)<1e-10)continue;
                        return false;
                    }
                    double expected=((Room)r.Element).Area*(top-bottom);
                    if(Math.Abs(solid.Volume-expected)>Math.Max(1e-7,expected*1e-7))return false;
                    Notice("ROOM_PRISM_AREA","Информация","Постоянство сечения подтверждено геометрией Room и балансом объёма. Используется Room.Area; наличие отдельного элемента чистого пола не требуется.",r);
                    return true;
                }
                catch(OperationCanceledException){throw;}
                catch{return false;} // An unverified room retains the exact section and floor checks.
            }
            public static bool IsRoomAreaMaskRole(string role)
            {return OneOf(role,"multilight","stair-gap","opening","shaft","engineering-shaft","stove","decoration");}
            private bool RoomAreaMaskRequired(Record r,Indicator metric)
            {
                if(r.Override==false)return true;
                var include=Mapped(r.Element,"include");
                return Eq(include,"0")||Eq(include,"нет")||Eq(include,"false")||IsRoomAreaMaskRole(r.Role);
            }
            private bool RoomHasVariableWalls(Record r)
            {
                var boundaries = ((Room)r.Element).GetBoundarySegments(new SpatialElementBoundaryOptions { SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish });
                if (boundaries == null) return true;
                foreach (var segment in boundaries.SelectMany(x => x))
                {
                    Element element = r.Element.Document.GetElement(segment.ElementId);
                    var link = element as RevitLinkInstance;
                    if (link != null) element = link.GetLinkDocument()?.GetElement(segment.LinkElementId);
                    if (element is ModelCurve) continue;
                    var wall = element as Wall;
                    if (wall == null || wall.IsStackedWall || wall.WallType.Kind != WallKind.Basic) return true;
                    BuiltInParameter section;
                    if (Enum.TryParse("WALL_CROSS_SECTION", out section) && (wall.get_Parameter(section)?.AsInteger() ?? 0) != 0) return true;
                    foreach (var side in new[] { ShellLayerType.Exterior, ShellLayerType.Interior })
                        foreach (var reference in HostObjectUtils.GetSideFaces(wall, side))
                        {
                            var face = wall.GetGeometryObjectFromReference(reference) as Face;
                            var plane = face as PlanarFace;
                            var cylinder = face as CylindricalFace;
                            if (plane != null && Math.Abs(plane.FaceNormal.Z) < 1e-10) continue;
                            if (cylinder != null && Math.Abs(Math.Abs(cylinder.Axis.Z) - 1) < 1e-10) continue;
                            return true;
                        }
                }
                return false;
            }

            private void CalculateRoomAreas(List<Record> records, Indicator metric)
            {
                var rooms=records.Where(r=>r.Element is Room&&r.Level!=null&&r.Level.Include&&r.Source.Mode!="reference").ToList();
                var includedRooms=new List<Record>();
                foreach(var r in rooms)
                {
                    try{if(Eligible(r,metric))includedRooms.Add(r);}
                    catch(OperationCanceledException){throw;}
                    catch(Exception ex){Notice("ROOM_AREA_INPUT","Ошибка",ex.Message,r,metric.ToString());}
                }
                // There is nothing to subtract from when the metric has no candidate rooms.
                if(includedRooms.Count==0)return;
                var targetFloors=new HashSet<string>(includedRooms.Select(r=>r.Building+"|"+Math.Round(r.Z,6)));
                // Keep explicit modelled openings/obstructions on the established exclusion path.
                // It applies the same method rules and planar masks, including non-Room geometry.
                if (records.Any(r => !(r.Element is Room) && r.Level != null && r.Level.Include &&
                    OneOf(r.Role, "multilight", "stair-gap", "opening", "shaft", "engineering-shaft", "stove", "decoration")))
                {
                    current.Issue("ROOM_SPECIAL_MASKS", "Информация", "Есть отдельные элементы нормативных исключений. Для этого показателя сохранён расчёт по контурам и маскам исключений.", metric.ToString());
                    CalculateAreas(records, metric);
                    return;
                }
                var indices = new Dictionary<string, PlanIndex>();
                var masks = new Dictionary<string, PlanIndex>();
                var rows = new List<Detail>();
                var invalidMaskFloors = new HashSet<string>();
                var reportedMaskFloors = new HashSet<string>();
                foreach (var r in rooms.Where(r=>targetFloors.Contains(r.Building+"|"+Math.Round(r.Z,6))))
                {
                    try
                    {
                        if (Eligible(r, metric)||!RoomAreaMaskRequired(r,metric)) continue;
                        Progress("Маски исключённых помещений: " + metric + "; ID " + IDHelper.ElIdValue(r.Element.Id));
                        var mask = RoomRequiresHeightSection(r) ? Plan(r, metric) : RoomAreaBoundary(r);
                        if (mask == null) continue;
                        string key = r.Building + "|" + Math.Round(r.Z, 6);
                        PlanIndex list;
                        if (!masks.TryGetValue(key, out list)) masks[key] = list = new PlanIndex();
                        list.Add(mask, r);
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (Exception ex)
                    {
                        invalidMaskFloors.Add(r.Building + "|" + Math.Round(r.Z, 6));
                        Notice("ROOM_AREA_MASK", "Ошибка", ex.Message, r, metric.ToString());
                    }
                }
                foreach (var r in includedRooms)
                {
                    Progress("Площади помещений: " + MetricLabels[metric.ToString()] + "; ID " + IDHelper.ElIdValue(r.Element.Id));
                    try
                    {
                        bool included = Eligible(r, metric) && r.Source.Mode != "reference";
                        if (!included) continue;
                        string key = r.Building + "|" + Math.Round(r.Z, 6);
                        if(invalidMaskFloors.Contains(key))
                        {
                            if(reportedMaskFloors.Add(key))Notice("ROOM_AREA_FLOOR_SKIPPED","Ошибка","Этаж «"+r.Level.Name+"» пропущен для площади помещений: не удалось проверить маску исключения. Остальные этажи рассчитываются.",r,metric.ToString());
                            continue;
                        }
                        var room = (Room)r.Element;
                        double raw = room.Area * .09290304;
                        bool special = RoomRequiresHeightSection(r);
                        Solid shape = special ? Plan(r, metric) : RoomAreaBoundary(r);
                        if (shape == null) continue;
                        bool excluded = VerticalExclusion(r, metric, records, shape);
                        // Keep the original niche/arch tests without constructing a room volume.
                        bool residential = metric == Indicator.ApartmentsTotal || metric == Indicator.ApartmentsHeated || !IsPublic(r) && r.Part != "nonresidential" && r.Role != "public";
                        if (residential && r.Role == "niche" && Dimension(r, "height", true) < 2) excluded = true;
                        if (residential && r.Role == "arch" && Dimension(r, "width", true) < 2) excluded = true;
                        if (special) raw = shape.Volume * .09290304;
                        else
                        {
                            double tolerance = Math.Max(1e-7, RequireLayeredBody(shape).Layers.Single().Region.Perimeter * ArcChordTolerance * 2) * .09290304;
                            if (Math.Abs(raw - shape.Volume * .09290304) > tolerance)
                                throw new InvalidOperationException("Площадь Room отличается от чистового контура. Проверьте способ вычисления границ помещений в Revit; значение не подменено автоматически.");
                        }
                        PlanIndex previous;
                        if (!excluded)
                        {
                            PlanIndex exclusions;
                            double initial = shape.Volume;
                            if (masks.TryGetValue(key, out exclusions))
                                foreach (var mask in exclusions.Query(shape))
                                {
                                    shape = Subtract(shape, mask);
                                    if (shape == null) break;
                                }
                            raw = Math.Max(0, raw - (initial - (shape?.Volume ?? 0)) * .09290304);
                            if (shape == null || raw <= 1e-9) continue;
                            if (!indices.TryGetValue(key, out previous)) indices[key] = previous = new PlanIndex();
                            foreach (var other in previous.Query(shape))
                                if ((Intersect(shape, other)?.Volume ?? 0) * .09290304 > .005)
                                    throw new InvalidOperationException("Помещение пересекается с ранее учтённым более чем на 0,005 м² и не включено в сумму. Проверьте дубликаты помещений и связей. Результат неполный.");
                            previous.Add(shape, r);
                        }
                        var row = Row(r, metric, raw, Factor(r, metric), excluded,
                            special ? "Нормативное сечение / высотное исключение" : "Площадь Room.Area; коэффициент применяется один раз", shape);
                        if (!special)
                        {
                            row.MeasurementElevation = r.Source.Transform.OfPoint(new XYZ(0, 0, room.Level.ProjectElevation + room.Level.get_Parameter(BuiltInParameter.LEVEL_ROOM_COMPUTATION_HEIGHT).AsDouble())).Z;
                            row.Reason = "Площадь Room.Area; чистовая граница Revit; коэффициент применяется один раз" + (excluded ? "; исключено нормативной проверкой" : "");
                        }
                        rows.Add(row);
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (Exception ex) { Notice("ROOM_AREA", "Ошибка", ex.Message, r, metric.ToString()); }
                }
                // Publish available rooms; metric-scoped errors mark the sum incomplete.
                // Failed exclusion masks invalidate their whole floor, never silently inflate its area.
                current.Details.AddRange(rows);
            }
        }
    }
}
