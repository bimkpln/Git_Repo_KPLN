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
            // Revit internal feet. This is a coincidence tolerance, never an elevation offset.
            private const double FloorGeometryTolerance = 1e-6;
            private readonly Dictionary<Tuple<Wall, ShellLayerType>, IList<Face>> floorWallFaces =
                new Dictionary<Tuple<Wall, ShellLayerType>, IList<Face>>();
            private readonly Dictionary<Face, Tuple<double, double>> floorFaceHeights = new Dictionary<Face, Tuple<double, double>>();

            private static void RequireHorizontalTransform(Transform transform)
            {
                if (Math.Abs(transform.BasisZ.Z - 1) > 1e-10 || Math.Abs(transform.Scale - 1) > 1e-10)
                    throw new InvalidOperationException("Наклонённая или масштабированная связь не задаёт горизонтальную плоскость пола. Источник не рассчитан.");
            }

            private double RoomFloorElevation(Record record)
            {
                if(IsClassifiedFamily(record.Element))
                {
                    if(record.Level==null)throw new InvalidOperationException("У экземпляра не определён расчётный уровень.");
                    return record.Z;
                }
                var room = record.Element as Room;
                if (room?.Level == null) throw new InvalidOperationException("Не определён уровень помещения.");
                RequireHorizontalTransform(record.Source.Transform);
                double offset = room.BaseOffset;
                // ProjectElevation is independent of the level's display elevation base.
                return record.Source.Transform.OfPoint(new XYZ(0, 0, room.Level.ProjectElevation + offset)).Z;
            }

            private static double RoomMeasurementHeight(string boundary, string metric, string role)
            {
                // Original specification, residential A.2.2: apartment room dimensions at 1.1-1.3 m.
                // Use the lower permitted height for heated apartment rooms; floor envelopes stay at 0.
                return boundary == "net" && OneOf(metric, "ApartmentsTotal", "ApartmentsHeated") &&
                    OneOf(role, "heated", "auxiliary", "mezzanine", "niche", "arch", "under-stair") ? 1.1 : 0;
            }

            private string RoomPlanCacheKey(Record record, string boundary, string metric)
            {
                return record.Key + "/" + boundary + "/floor=" + RoomFloorElevation(record).ToString("R", CultureInfo.InvariantCulture) +
                    "/height=" + RoomMeasurementHeight(boundary, metric, record.Role).ToString("R", CultureInfo.InvariantCulture) + "/role=" + record.Role;
            }

            private static bool CrossesFloor(Element element, Transform transform, double elevation)
            {
                var box = element.get_BoundingBox(null);
                if (box == null) return false;
                var placement = transform.Multiply(box.Transform);
                var z = Enumerable.Range(0, 8).Select(i => placement.OfPoint(new XYZ(
                    (i & 1) == 0 ? box.Min.X : box.Max.X,
                    (i & 2) == 0 ? box.Min.Y : box.Max.Y,
                    (i & 4) == 0 ? box.Min.Z : box.Max.Z)).Z).ToList();
                // At a storey joint select the wall starting here, not the wall ending here.
                return elevation >= z.Min() - FloorGeometryTolerance && elevation < z.Max() - FloorGeometryTolerance;
            }

            private static IEnumerable<Wall> WallMembers(Wall wall)
            {
                return wall.IsStackedWall ? wall.GetStackedWallMemberIds().Select(id => wall.Document.GetElement(id)).OfType<Wall>() : new[] { wall };
            }

            private sealed class FloorWallTrace
            {
                internal double Offset;
                internal double Outward; // Face normal projected on the curve's horizontal right normal.
                internal double Slope;
            }

            // Plane equation n.(p-origin)=0 evaluated at the requested Z. No sampled cut height.
            private static double PlaneSectionOffset(double signedDistance, double normalZ, double elevationDelta, double normalRight)
            {
                if (Math.Abs(normalRight) < 1e-10) throw new InvalidOperationException("Поверхность не задаёт боковую границу в плоскости пола.");
                return (signedDistance - normalZ * elevationDelta) / normalRight;
            }

            private static double ConeSectionRadius(double radius, double radialNormal, double normalZ, double elevationDelta)
            {
                if (Math.Abs(radialNormal) < 1e-10) throw new InvalidOperationException("Не определена образующая конической стены.");
                double result = radius - normalZ / radialNormal * elevationDelta;
                if (result <= FloorGeometryTolerance) throw new InvalidOperationException("Сечение конической стены вырождено.");
                return result;
            }

            private FloorWallTrace TraceWallFace(Face face, Transform transform, Curve curve, double elevation)
            {
                RequireHorizontalTransform(transform);
                Tuple<double, double> range;
                if (!floorFaceHeights.TryGetValue(face, out range))
                {
                    var vertices = face.Triangulate().Vertices;
                    if (vertices.Count == 0) return null;
                    floorFaceHeights[face] = range = Tuple.Create(vertices.Min(p => p.Z), vertices.Max(p => p.Z));
                }
                double localElevation = elevation - transform.Origin.Z;
                if (localElevation < range.Item1 - FloorGeometryTolerance || localElevation >= range.Item2 - FloorGeometryTolerance) return null;
                var point = curve.Evaluate(.5, true);
                var right = curve.ComputeDerivatives(.5, true).BasisX.Normalize().CrossProduct(XYZ.BasisZ).Normalize();
                var plane = face as PlanarFace;
                if (plane != null)
                {
                    var normal = transform.OfVector(plane.FaceNormal).Normalize();
                    double horizontal = Math.Sqrt(normal.X * normal.X + normal.Y * normal.Y);
                    if (!(curve is Line) || horizontal < 1e-10 || Math.Abs(Math.Abs(normal.DotProduct(right)) - horizontal) > 1e-8) return null;
                    var origin = transform.OfPoint(plane.Origin);
                    return new FloorWallTrace
                    {
                        Offset = PlaneSectionOffset((origin - point).DotProduct(normal), normal.Z, elevation - point.Z, normal.DotProduct(right)),
                        Outward = normal.DotProduct(right),
                        Slope = -normal.Z / normal.DotProduct(right)
                    };
                }
                var arc = curve as Arc;
                var cylinder = face as CylindricalFace;
                var cone = face as ConicalFace;
                if (arc == null || cylinder == null && cone == null) return null;
                var axis = transform.OfVector(cylinder != null ? cylinder.Axis : cone.Axis).Normalize();
                var center = transform.OfPoint(cylinder != null ? cylinder.Origin : cone.Origin);
                if (Math.Abs(Math.Abs(axis.Z) - 1) > 1e-10 ||
                    new XYZ(center.X, center.Y, 0).DistanceTo(new XYZ(arc.Center.X, arc.Center.Y, 0)) > FloorGeometryTolerance) return null;
                var uvBox = face.GetBoundingBox();
                var uv = new UV((uvBox.Min.U + uvBox.Max.U) / 2, (uvBox.Min.V + uvBox.Max.V) / 2);
                var sample = transform.OfPoint(face.Evaluate(uv));
                var n = transform.OfVector(face.ComputeNormal(uv)).Normalize();
                var radialSample = new XYZ(sample.X - center.X, sample.Y - center.Y, 0);
                double radialNormal = n.DotProduct(radialSample.Normalize());
                double radius = ConeSectionRadius(radialSample.GetLength(), radialNormal, n.Z, elevation - sample.Z);
                var radial = new XYZ(point.X - center.X, point.Y - center.Y, 0).Normalize();
                return new FloorWallTrace { Offset = (radius - arc.Radius) / right.DotProduct(radial), Outward = radialNormal * radial.DotProduct(right),
                    Slope = -n.Z / radialNormal / right.DotProduct(radial) };
            }

            private FloorWallTrace WallFaceAtFloor(Wall wall, Transform transform, Curve curve, double elevation, ShellLayerType side)
            {
                var key = Tuple.Create(wall, side);
                IList<Face> faces;
                if (!floorWallFaces.TryGetValue(key, out faces))
                    floorWallFaces[key] = faces = HostObjectUtils.GetSideFaces(wall, side).Select(wall.GetGeometryObjectFromReference).OfType<Face>().ToList();
                var traces = faces.Select(f => TraceWallFace(f, transform, curve, elevation)).Where(t => t != null).ToList();
                if (traces.Count == 0 || traces.Max(t => t.Offset) - traces.Min(t => t.Offset) > FloorGeometryTolerance ||
                    traces.Any(t => t.Outward * traces[0].Outward <= 0 || Math.Abs(t.Slope - traces[0].Slope) > 1e-8))
                    throw new InvalidOperationException("Стена ID " + IDHelper.ElIdValue(wall.Id) + ": не получено однозначное сечение боковой поверхности на отметке пола " +
                        (elevation * .3048).ToString("0.######", CultureInfo.InvariantCulture) + " м. Поддерживаются плоские, цилиндрические и соосные конические поверхности; неоднозначная геометрия не заменяется осью стены.");
                return traces[0];
            }

            private double WallCoreAtFloor(Wall wall, FloorWallTrace exterior, FloorWallTrace interior, bool exteriorSide)
            {
                var structure = wall.WallType.GetCompoundStructure();
                if (structure == null || structure.GetFirstCoreLayerIndex() < 0 || structure.IsVerticallyCompound)
                    throw new InvalidOperationException("Стена ID " + IDHelper.ElIdValue(wall.Id) + ": ядро не определено либо имеет вертикально составную структуру. Граница ядра не подменена наружной отделкой.");
                var layers = structure.GetLayers();
                double finish = exteriorSide ? layers.Take(structure.GetFirstCoreLayerIndex()).Sum(l => l.Width) :
                    layers.Skip(structure.GetLastCoreLayerIndex() + 1).Sum(l => l.Width);
                var selected = exteriorSide ? exterior : interior;
                if (finish < 1e-10) return selected.Offset;
                if (Math.Abs(exterior.Slope - interior.Slope) > 1e-8)
                    throw new InvalidOperationException("Стена ID " + IDHelper.ElIdValue(wall.Id) + ": боковые поверхности имеют разный наклон. Граница ядра не определяется пропорцией номинальных слоёв.");
                BuiltInParameter parameter;
                if (Enum.TryParse("WALL_CROSS_SECTION", out parameter) && (wall.get_Parameter(parameter)?.AsInteger() ?? 0) == 2)
                    throw new InvalidOperationException("Стена ID " + IDHelper.ElIdValue(wall.Id) + ": переменное сечение с отделочными слоями. Для ГНС нужна фактическая граница ядра; средняя толщина не используется.");
                double width = layers.Sum(l => l.Width);
                if (width <= 0) throw new InvalidOperationException("Нулевая толщина структуры стены.");
                // Uniform compound layers share the same inclination. Use measured horizontal width,
                // not nominal wall.Width / 2; this also handles a shifted location line.
                return selected.Offset + ((exteriorSide ? interior : exterior).Offset - selected.Offset) * finish / width;
            }

            private Solid RoomPlanAtFloor(Record record, string boundary, string metric)
            {
                var room = (Room)record.Element;
                double floorElevation = RoomFloorElevation(record);
                double height = RoomMeasurementHeight(boundary, metric, record.Role);
                double elevation = floorElevation + height / .3048;
                if (boundary == "net")
                {
                    var native = BuildPlanarSolid(MeasuredRoomRegion(record, metric), 0, 1, true);
                    if (native == null) throw new InvalidOperationException("Revit не вернул замкнутое сечение помещения на отметке обмера.");
                    Notice("ROOM_NATIVE_SECTION", "Информация", "Чистая площадь получена сечением геометрии Room на отметке " + (elevation * .3048).ToString("0.######", CultureInfo.InvariantCulture) + " м. Типы образующих границу элементов не ограничиваются стенами.", record, metric);
                    return native;
                }
                RoomBoundaryPlan(record, "net");
                if (boundary == "gross")
                {
                    var centrePlan = Section(SpatialVolume(record, true), floorElevation);
                    if (centrePlan == null) throw new InvalidOperationException("Нет сечения помещения по осям внутренних ограждений.");
                    Solid grossPlan = null;
                    foreach (var envelope in RoomFloorContours(new List<Record> { record }, record, Indicator.Gross))
                        grossPlan = Union(grossPlan, Intersect(centrePlan, envelope));
                    if (grossPlan == null) throw new InvalidOperationException("Помещение не пересекло внутренний контур этажа.");
                    return grossPlan;
                }
                var options = new SpatialElementBoundaryOptions { SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Center };
                var boundaries = room.GetBoundarySegments(options);
                if (boundaries == null || boundaries.Count == 0) throw new InvalidOperationException("Нет замкнутых границ помещения.");
                var loops = boundaries.Select(b => FlatLoop(b.Select(s => s.GetCurve().CreateTransformed(record.Source.Transform)))).ToList();
                var centre = Extrude(loops);
                var result = new List<CurveLoop>();
                for (int li = 0; li < loops.Count; li++)
                {
                    var offsets = new List<double>();
                    for (int si = 0; si < boundaries[li].Count; si++)
                    {
                        var segment = boundaries[li][si];
                        var element = room.Document.GetElement(segment.ElementId);
                        var transform = record.Source.Transform;
                        var link = element as RevitLinkInstance;
                        if (link != null)
                        {
                            element = link.GetLinkDocument()?.GetElement(segment.LinkElementId);
                            transform = transform.Multiply(link.GetTotalTransform());
                        }
                        var curve = loops[li].ElementAt(si);
                        var wall = element as Wall;
                        if (wall == null)
                        {
                            if (element is ModelCurve && IDHelper.ElIdValue(element.Category.Id) == (long)BuiltInCategory.OST_RoomSeparationLines)
                            { offsets.Add(0); continue; }
                            throw new InvalidOperationException("Граница помещения не связана с доступной стеной или разделителем помещений; ID границы " +
                                IDHelper.ElIdValue(segment.ElementId) + ". Её положение на отметке пола не подтверждено.");
                        }
                        if (wall.IsStackedWall)
                        {
                            var members = WallMembers(wall).Where(w => CrossesFloor(w, transform, elevation)).ToList();
                            if (members.Count != 1) throw new InvalidOperationException("Составная стена ID " + IDHelper.ElIdValue(wall.Id) + ": на отметке обмера не определён единственный участок.");
                            wall = members[0];
                        }
                        var outer = WallFaceAtFloor(wall, transform, curve, elevation, ShellLayerType.Exterior);
                        var inner = WallFaceAtFloor(wall, transform, curve, elevation, ShellLayerType.Interior);
                        var mid = curve.Evaluate(.5, true);
                        var right = curve.ComputeDerivatives(.5, true).BasisX.Normalize().CrossProduct(XYZ.BasisZ).Normalize();
                        // Probe only determines which side of the closed ring is inside, never a cut elevation.
                        bool rightInside = Inside(centre, mid + right * .002), leftInside = Inside(centre, mid - right * .002);
                        if (rightInside == leftInside) throw new InvalidOperationException("Не определена внутренняя сторона границы помещения. Узкий или пересекающийся участок не исключён из площади молча.");
                        double sign = rightInside ? -1 : 1;
                        bool exteriorIsAway = outer.Outward * sign > 0;
                        var facingRoom = exteriorIsAway ? inner : outer;
                        double offset;
                        if (boundary == "net") offset = facingRoom.Offset;
                        else if (wall.WallType.Function != WallFunction.Exterior) offset = (outer.Offset + inner.Offset) / 2;
                        else if (boundary == "gross") offset = facingRoom.Offset;
                        else if (Methodology.UsesCoreBoundary) offset = WallCoreAtFloor(wall, outer, inner, exteriorIsAway);
                        else offset = (exteriorIsAway ? outer : inner).Offset;
                        offsets.Add(offset);
                    }
                    var shifted = CurveLoop.CreateViaOffset(loops[li], offsets, XYZ.BasisZ);
                    // Revit may drop a short segment while offsetting. Never accept that silently.
                    if (shifted.Count() != loops[li].Count()) throw new InvalidOperationException("На отметке пола изменилось число границ помещения. Требуется разбор фактической топологии; сегменты не удалены из расчёта.");
                    result.Add(shifted);
                }
                var plan = Extrude(result);
                Notice("ROOM_FLOOR_DATUM", "Информация", "Пол: " + (floorElevation * .3048).ToString("0.######", CultureInfo.InvariantCulture) +
                    " м; " + (string.IsNullOrWhiteSpace(Config.Parameter("floor-offset")) ? "низ Room; отдельное перекрытие не требуется" : "параметр «" + Config.Parameter("floor-offset") + "»") + ".", record, metric);
                if (height > 0) Notice("ROOM_MEASUREMENT_HEIGHT", "Информация", "Обмер помещения квартиры на высоте 1,1 м от пола: нижняя граница диапазона 1,1-1,3 м по А.2.2 исходного ТЗ. Поэтажные контуры остаются на уровне пола.", record, metric);
                return plan;
            }
        }
    }
}
