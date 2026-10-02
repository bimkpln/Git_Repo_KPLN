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
                return roomHeightSections[r.Key] = RoomHasVariableWalls(r);
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
                int firstIssue = current.Issues.Count;
                foreach (var r in records.Where(r => r.Element is Room && r.Level != null && r.Level.Include && r.Source.Mode != "reference"))
                {
                    try
                    {
                        if (Eligible(r, metric)) continue;
                        Progress("Маски исключённых помещений: " + metric + "; ID " + IDHelper.ElIdValue(r.Element.Id));
                        var mask = RoomRequiresHeightSection(r) ? Plan(r, metric) : RoomAreaBoundary(r);
                        if (mask == null) continue;
                        string key = r.Building + "|" + Math.Round(r.Z, 6);
                        PlanIndex list;
                        if (!masks.TryGetValue(key, out list)) masks[key] = list = new PlanIndex();
                        list.Add(mask, r);
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (Exception ex) { Notice("ROOM_AREA_MASK", "Ошибка", ex.Message, r, metric.ToString()); }
                }
                foreach (var r in records.Where(r => r.Element is Room && r.Level != null && r.Level.Include))
                {
                    Progress("Площади помещений: " + MetricLabels[metric.ToString()] + "; ID " + IDHelper.ElIdValue(r.Element.Id));
                    try
                    {
                        bool included = Eligible(r, metric) && r.Source.Mode != "reference";
                        if (!included) continue;
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
                        string key = r.Building + "|" + Math.Round(r.Z, 6);
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
                                    throw new InvalidOperationException("Учтённые помещения пересекаются более чем на 0,005 м². Проверьте дубликаты помещений и связей. Сумма показателя не опубликована.");
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
                // No plausible-looking partial apartment totals after a failed room.
                if (current.Issues.Skip(firstIssue).Any(i => i.Severity == "Ошибка")) return;
                current.Details.AddRange(rows);
            }
        }
    }
}
