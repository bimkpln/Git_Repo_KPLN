using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Linq;
using CT = KPLN_CalculateTEP.Common.Geometry.TepClipper.ClipType;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public partial class Engine
        {
            private static PlanarRegion RegionUnion(PlanarRegion a, PlanarRegion b)
            { return PlanarRegion.Combine(a, b, CT.ctUnion); }
            private static PlanarRegion RegionIntersection(PlanarRegion a, PlanarRegion b)
            { return PlanarRegion.Combine(a, b, CT.ctIntersection); }
            private static PlanarRegion RegionDifference(PlanarRegion a, PlanarRegion b)
            { return PlanarRegion.Combine(a, b, CT.ctDifference); }

            // A numeric result never needs to survive a second extrusion through the Revit kernel.
            private static PlanarRegion RegionFromLoops(IEnumerable<CurveLoop> input, bool validateBoundary=false)
            {
                var loops = input.ToList();
                if (loops.Count == 0) return PlanarRegion.Empty;
                double tolerance = ArcChordTolerance;
                for (int attempt = 0; attempt < 8; attempt++, tolerance /= 4)
                {
                    bool curved = false;
                    var rings = loops.Select(l => PolygonRing(l, ref curved, tolerance)).ToList();
                    var result = PlanarRegion.FromRings(rings);
                    // Convex-hull chord bound; sum ring perimeters BEFORE union so cancellation
                    // or overlapping rings cannot conceal an approximation error.
                    double perimeter = rings.Sum(r => r.Select((p, i) => {
                        var q = r[(i + 1) % r.Count];
                        return Math.Sqrt((p[0]-q[0])*(p[0]-q[0])+(p[1]-q[1])*(p[1]-q[1]));
                    }).Sum());
                    if (!curved || (2 * perimeter * tolerance + loops.Count * Math.PI * tolerance * tolerance) * .09290304 <= .001)
                    {
                        if(validateBoundary)PlanarRegion.ValidateBoundaryRings(rings.Select(r=>r.ToArray()).ToList());
                        return result;
                    }
                }
                throw new InvalidOperationException("Не подтверждён допуск аппроксимации площади кривых 0,001 м². Точность не снижена.");
            }

            private PlanarRegion NativeBoundaryRegion(IList<Curve> curves,double elevation,IList<bool> calculationFrame=null)
            {
                double tolerance=ArcChordTolerance;
                for(int attempt=0;attempt<8;attempt++,tolerance/=4)
                {
                    var lines=new List<double[][]>();double perimeter=0;bool curved=false;
                    foreach(var curve in curves)
                    {
                        var points=CurvePoints(curve,tolerance);
                        if(points.Any(p=>Math.Abs(p[2]-elevation)>FloorGeometryTolerance))throw new UnconfirmedEnvelopeException("Граница Revit лежит не на расчётной отметке пола.");
                        for(int i=0;i+1<points.Count;i++){double dx=points[i+1][0]-points[i][0],dy=points[i+1][1]-points[i][1];perimeter+=Math.Sqrt(dx*dx+dy*dy);}
                        lines.Add(points.ToArray());curved|=!(curve is Line);
                    }
                    if(curved&&(2*perimeter*tolerance+curves.Count*Math.PI*tolerance*tolerance)*.09290304>.001)continue;
                    return calculationFrame==null?PlanarRegion.FromBoundarySegments(lines,planarCheckpoint):
                        PlanarRegion.FromExteriorBoundarySegments(lines,calculationFrame,planarCheckpoint);
                }
                throw new UnconfirmedEnvelopeException("Не подтверждён допуск площади нативной границы 0,001 м².");
            }

            private static PlanarRegion BoundaryBands(IList<Curve> curves,IList<double> offsets)
            {
                if(curves.All(c=>c is Line))return PlanarRegion.StraightBoundaryBands(curves.Select(c=>{var p=c.GetEndPoint(0);return new[]{p.X,p.Y};}).ToArray(),offsets.ToArray());
                var shifted=new List<Curve>();
                for(int i=0;i<curves.Count;i++)
                {
                    if(Math.Abs(offsets[i])<1e-12){shifted.Add(curves[i]);continue;}
                    // Revit's ellipse/spline offset is an approximation with no area error bound.
                    if(!(curves[i] is Line)&&!(curves[i] is Arc))
                        throw new UnconfirmedEnvelopeException("Для внутренней стороны кривой "+curves[i].GetType().Name+" нет точного параллельного представления; приближённое смещение не использовано.");
                    var curve=curves[i].CreateOffset(offsets[i],XYZ.BasisZ);
                    if(curve==null)throw new UnconfirmedEnvelopeException("Не определена внутренняя сторона участка ограждения.");
                    shifted.Add(curve);
                }
                double tolerance=ArcChordTolerance;
                for(int attempt=0;attempt<8;attempt++,tolerance/=4)
                {
                    var bands=PlanarRegion.Empty;double perimeter=0;bool curved=false;
                    for(int i=0;i<curves.Count;i++)
                    {
                        if(Math.Abs(offsets[i])<1e-12)continue;
                        var a=CurvePoints(curves[i],tolerance);var b=CurvePoints(shifted[i],tolerance);
                        var polygon=a.Concat(b.AsEnumerable().Reverse()).ToArray();
                        var band=PlanarRegion.FromRings(new[]{polygon});
                        perimeter+=band.Perimeter;curved|=!(curves[i] is Line);
                        bands=RegionUnion(bands,band);
                    }
                    if(curved&&(2*perimeter*tolerance+curves.Count*Math.PI*tolerance*tolerance)*.09290304>.001)continue;
                    for(int i=0;i<curves.Count;i++)
                    {
                        int next=(i+1)%curves.Count;var a=shifted[i].GetEndPoint(1);var b=shifted[next].GetEndPoint(0);
                        if(a.DistanceTo(b)<=2/PlanarRegion.Scale)continue;
                        if(!(curves[i] is Line)||!(curves[next] is Line))
                            throw new UnconfirmedEnvelopeException("Несовпадающие внутренние границы криволинейного стыка требуют точного пересечения; соединение не подменено прямой.");
                        var corner=curves[i].GetEndPoint(1);var da=shifted[i].ComputeDerivatives(1,true).BasisX;var db=shifted[next].ComputeDerivatives(0,true).BasisX;
                        bands=RegionUnion(bands,PlanarRegion.BoundaryBandJoin(new[]{corner.X,corner.Y},new[]{a.X,a.Y},new[]{da.X,da.Y},new[]{b.X,b.Y},new[]{db.X,db.Y}));
                    }
                    return bands;
                }
                throw new UnconfirmedEnvelopeException("Не подтверждён допуск площади полос ограждений 0,001 м².");
            }

            private static PlanarRegion RegionOfPlan(Solid plan)
            {
                if (plan == null) return PlanarRegion.Empty;
                var body = RequireLayeredBody(plan);
                if (body.Layers.Count != 1) throw new InvalidOperationException("Вместо плоской области получено тело с несколькими высотными участками.");
                return body.Layers[0].Region;
            }

            private static PlanarRegion SectionRegion(Solid solid, double z)
            {
                if (solid == null) return PlanarRegion.Empty;
                LayeredBody body; string reason;
                if (TryLayeredBody(solid, out body, out reason))
                {
                    var region = PlanarRegion.Empty;
                    foreach (var layer in body.Layers.Where(l => l.Bottom <= z + 1e-8 && l.Top >= z - 1e-8))
                        region = RegionUnion(region, layer.Region);
                    return region;
                }
                var bounds = Bounds(solid);
                if (z < bounds[2] - 1e-8 || z > bounds[5] + 1e-8) return PlanarRegion.Empty;
                var boundary = PlanarRegion.Empty;
                foreach (var face in solid.Faces.Cast<Face>().OfType<PlanarFace>().Where(f => Math.Abs(Math.Abs(f.FaceNormal.Z)-1)<1e-10 && Math.Abs(f.Origin.Z-z)<1e-7))
                    boundary = RegionUnion(boundary, RegionFromLoops(face.GetEdgesAsCurveLoops()));
                if (z <= bounds[2] || z >= bounds[5]) return boundary;
                var origin = new XYZ((bounds[0]+bounds[3])/2, (bounds[1]+bounds[4])/2, z);
                var local = SolidUtils.CreateTransformed(solid, Transform.CreateTranslation(-origin));
                var failures = new List<string>();
                foreach (bool above in new[] { false, true })
                {
                    try
                    {
                        planarCheckpoint?.Invoke();
                        var cut = BooleanOperationsUtils.CutWithHalfSpace(local, Plane.CreateByNormalAndOrigin(above ? XYZ.BasisZ : -XYZ.BasisZ, XYZ.Zero));
                        var result = boundary; bool found = false;
                        if (cut != null) foreach (var face in cut.Faces.Cast<Face>().OfType<PlanarFace>().Where(f => Math.Abs(Math.Abs(f.FaceNormal.Z)-1)<1e-10 && Math.Abs(f.Origin.Z)<1e-7))
                        {
                            result = RegionUnion(result, RegionFromLoops(face.GetEdgesAsCurveLoops().Select(l => CurveLoop.Create(l.Select(c => c.CreateTransformed(Transform.CreateTranslation(origin))).ToList()))));
                            found = true;
                        }
                        if (found) return result;
                        failures.Add("Не получена грань сечения.");
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (Exception ex) { failures.Add(ex.Message); }
                }
                throw new InvalidOperationException("Не получено сечение на заданной отметке без смещения высоты: " + string.Join(" | ", failures));
            }

            private static PlanarRegion ProjectionRegion(Solid solid)
            {
                if (solid == null) return PlanarRegion.Empty;
                LayeredBody body; string reason;
                if (TryLayeredBody(solid, out body, out reason)) return body.Projection();
                var result = PlanarRegion.Empty;
                foreach (Face face in solid.Faces)
                {
                    var plane = face as PlanarFace;
                    if (plane == null)
                    {
                        var cylinder = face as CylindricalFace;
                        if (cylinder != null && Math.Abs(cylinder.Axis.Z) > .999999) continue;
                        throw new InvalidOperationException("Для криволинейной поверхности не подтверждена точная горизонтальная проекция.");
                    }
                    if (plane.FaceNormal.Z > 1e-7)
                        result = RegionUnion(result, RegionFromLoops(plane.GetEdgesAsCurveLoops().Select(l => FlatLoop(l))));
                }
                if (result.IsEmpty) throw new InvalidOperationException("Не найдена горизонтальная проекция тела.");
                return result;
            }

            private static PlanarRegion FootprintRegion(Solid solid, double lowest, double groundElevation)
            {
                if (lowest > groundElevation) return PlanarRegion.Empty;
                LayeredBody body; string reason;
                if (TryLayeredBody(solid, out body, out reason)) return body.Footprint(lowest, groundElevation);
                var clipped = Half(solid, lowest, true);
                return RegionUnion(SectionRegion(clipped, groundElevation), ProjectionRegion(Half(clipped, groundElevation, false)));
            }

            private readonly Dictionary<string, PlanarRegion> measuredRoomRegions = new Dictionary<string, PlanarRegion>();
            private PlanarRegion MeasuredRoomRegion(Record record, string metric)
            {
                PlanarRegion edited;if(TryEditedObjectOverlay(record,out edited))return edited;
                if(IsClassifiedFamily(record.Element))return FamilyAreaRegion(record);
                string key=RoomPlanCacheKey(record,"net",metric);
                PlanarRegion result;
                if(measuredRoomRegions.TryGetValue(key,out result))return result;
                double elevation=RoomFloorElevation(record)+RoomMeasurementHeight("net",metric,record.Role)/.3048;
                result=ReviewRegion("room/"+key,record,"Помещение - чистая площадь",elevation,()=>SectionRegion(SpatialVolume(record),elevation));
                if(result.IsEmpty)throw new InvalidOperationException("Нет замкнутого сечения Room на нормативной отметке обмера "+(elevation*.3048).ToString("0.######",System.Globalization.CultureInfo.InvariantCulture)+" м.");
                measuredRoomRegions[key]=result;
                return result;
            }
            private PlanarRegion EligibleRoomRegion(Record r,Indicator metric)
            {
                return ApplyRoomHeightRules(r,metric,IsClassifiedFamily(r.Element)?FamilyAreaRegion(r):MeasuredRoomRegion(r,metric.ToString()),false);
            }
            private PlanarRegion ApplyRoomHeightRules(Record r,Indicator metric,PlanarRegion result,bool centre)
            {
                if((int)metric<=4)return result;
                bool residential=metric==Indicator.ApartmentsTotal||metric==Indicator.ApartmentsHeated||!IsPublic(r)&&r.Part!="nonresidential"&&r.Role!="public";
                if(residential&&r.Role=="arch"&&Dimension(r,"width",true)<2)return PlanarRegion.Empty;
                double minimum=residential&&r.Role=="under-stair"?1.600001:residential&&r.Role=="niche"?2:0;
                bool slope=NeedsCeilingCheck(r,metric)&&HasSlopedCeiling(r);
                if(minimum>0||slope)result=RegionIntersection(result,ClearanceMask(r,metric,minimum,slope,centre));
                return result;
            }

            private readonly Dictionary<string, PlanarRegion> roomNetRegions = new Dictionary<string, PlanarRegion>();
            private PlanarRegion RoomNetRegion(Record record)
            {
                string key = RoomPlanCacheKey(record, "net", "");
                PlanarRegion result;
                if (roomNetRegions.TryGetValue(key, out result)) return result;
                // Read the real section at the supplied room datum. A separate finish-floor
                // element is not a prerequisite for the room's geometric section.
                result = MeasuredRoomRegion(record, "");
                if (result.IsEmpty) throw new InvalidOperationException("Нет замкнутого сечения помещения на отметке пола.");
                roomNetRegions[key] = result;
                return result;
            }

            private Detail PlanarRow(Record record, Indicator metric, PlanarRegion region, bool excluded, string reason)
            {
                var row = Row(record, metric, region.Area * .09290304, 1, excluded, reason);
                if (record.Element is Room) row.MeasurementElevation = RoomFloorElevation(record);
                return row;
            }
        }
    }
}
