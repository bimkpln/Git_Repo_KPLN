using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        internal class HeightPatch
        {
            internal PlanarRegion Region;
            internal double X,Y,C;
            internal bool Top;
            internal bool RoofCeiling;
            internal double At(double x,double y){return X*x+Y*y+C;}
        }
        public partial class Engine
        {
            private static HeightPatch HeightFace(PlanarFace face,Transform transform)
            {
                var normal=transform.OfVector(face.FaceNormal);if(Math.Abs(normal.Z)<1e-10)return null;
                var origin=transform.OfPoint(face.Origin);bool curved=false;
                var rings=new List<List<double[]>>();
                foreach(var loop in face.GetEdgesAsCurveLoops())
                {
                    var transformed=new CurveLoop();foreach(var curve in loop)transformed.Append(curve.CreateTransformed(transform));
                    rings.Add(PolygonRing(transformed,ref curved));
                }
                // Curved boundaries on a sloping plane need an extremum calculation, not vertex sampling.
                if(curved&&(Math.Abs(normal.X)>1e-10||Math.Abs(normal.Y)>1e-10))
                    throw new InvalidOperationException("Наклонная грань имеет криволинейную границу; точные экстремумы высоты не подтверждены.");
                return new HeightPatch{Region=PlanarRegion.FromRings(rings),X=-normal.X/normal.Z,Y=-normal.Y/normal.Z,
                    C=origin.Z+(normal.X*origin.X+normal.Y*origin.Y)/normal.Z,Top=normal.Z>0};
            }
            private static double[] HeightRange(List<HeightPatch> patches)
            {
                var top=patches.Where(p=>p.Top).ToList();var bottom=patches.Where(p=>!p.Top).ToList();
                if(top.Count==0||bottom.Count==0)throw new InvalidOperationException("Не определены нижняя и верхняя поверхности.");
                foreach(var group in new[]{top,bottom})for(int i=0;i<group.Count;i++)for(int j=i+1;j<group.Count;j++)
                    if(RegionIntersection(group[i].Region,group[j].Region).Area*.09290304>.001)
                        throw new InvalidOperationException("Несколько вертикальных объёмов над одной точкой. Единую высоту определить нельзя.");
                var tops=top.Aggregate(PlanarRegion.Empty,(r,p)=>RegionUnion(r,p.Region));
                var bottoms=bottom.Aggregate(PlanarRegion.Empty,(r,p)=>RegionUnion(r,p.Region));
                if((RegionDifference(tops,bottoms).Area+RegionDifference(bottoms,tops).Area)*.09290304>.001)
                    throw new InvalidOperationException("Проекции нижней и верхней поверхностей не совпадают.");
                double min=double.PositiveInfinity,max=double.NegativeInfinity;
                foreach(var t in top)foreach(var b in bottom)
                    foreach(var p in RegionIntersection(t.Region,b.Region).Paths.SelectMany(p=>p))
                    {double h=t.At(p.X/PlanarRegion.Scale,p.Y/PlanarRegion.Scale)-b.At(p.X/PlanarRegion.Scale,p.Y/PlanarRegion.Scale);min=Math.Min(min,h);max=Math.Max(max,h);}
                if(double.IsInfinity(min)||min<=0)throw new InvalidOperationException("Нет положительной подтверждённой высоты.");
                return new[]{min*.3048,max*.3048};
            }
            private List<HeightPatch> HeightPatches(Solid solid,Transform transform)
            {
                var result=new List<HeightPatch>();
                foreach(Face face in solid.Faces)
                {
                    var plane=face as PlanarFace;
                    if(plane==null)
                    {
                        var cylinder=face as CylindricalFace;
                        if(cylinder!=null&&Math.Abs(Math.Abs(transform.OfVector(cylinder.Axis).Normalize().Z)-1)<1e-12)continue;
                        throw new InvalidOperationException("Криволинейная поверхность: точный диапазон высоты не подтверждён.");
                    }
                    var patch=HeightFace(plane,transform);if(patch!=null)result.Add(patch);
                }
                return result;
            }
            private void ConfirmRoomHeight(List<HeightPatch> patches,Record room,List<Tuple<Source,HostObject>> hosts,
                Dictionary<string,List<HeightPatch>> hostCache)
            {
                // Revit can cap a Room at its upper limit without a physical roof/ceiling.
                // Match its floor and ceiling to actual model faces, including linked structures.
                var bounds=ElementBounds(room.Element,room.Source.Transform);
                var supports=new List<HeightPatch>();
                foreach(var pair in hosts)
                {
                    var s=pair.Item1;var host=pair.Item2;var box=ElementBounds(host,s.Transform);
                    if(box==null||box[3]<bounds[0]||box[0]>bounds[3]||box[4]<bounds[1]||box[1]>bounds[4]||box[5]<bounds[2]-1e-6||box[2]>bounds[5]+1e-6)continue;
                    string key=s.Key+"/"+host.UniqueId;
                    List<HeightPatch> faces;
                    if(!hostCache.TryGetValue(key,out faces))
                    {
                        faces=new List<HeightPatch>();
                        foreach(var reference in HostObjectUtils.GetTopFaces(host).Concat(HostObjectUtils.GetBottomFaces(host)))
                        {
                            var face=host.GetGeometryObjectFromReference(reference) as PlanarFace;if(face==null)continue;
                            var patch=HeightFace(face,s.Transform);if(patch!=null)faces.Add(patch);
                        }
                        hostCache[key]=faces;
                    }
                    supports.AddRange(faces);
                }
                foreach(var patch in patches)
                {
                    var cover=PlanarRegion.Empty;
                    foreach(var support in supports.Where(p=>p.Top!=patch.Top&&Math.Abs(p.X-patch.X)<1e-10&&Math.Abs(p.Y-patch.Y)<1e-10&&Math.Abs(p.C-patch.C)<1e-6))
                        cover=RegionUnion(cover,support.Region);
                    if(RegionDifference(patch.Region,cover).Area*.09290304>.001)
                        throw new InvalidOperationException((patch.Top?"Верх":"Низ")+" помещения не подтверждён геометрией перекрытия, потолка или крыши. Граница по высоте помещения не подставлена как фактическая конструкция.");
                }
            }
        }
    }
}
