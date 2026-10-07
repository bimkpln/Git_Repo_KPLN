using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        internal class ClearanceGeometry
        {
            internal List<HeightPatch> Patches;
            internal double Minimum,Maximum,MinimumSlope,MaximumSlope;
        }
        public partial class Engine
        {
            private readonly Dictionary<string,ClearanceGeometry> clearanceGeometry=new Dictionary<string,ClearanceGeometry>();
            private readonly Dictionary<string,string> clearanceErrors=new Dictionary<string,string>();
            private readonly Dictionary<string,List<HeightPatch>> clearanceFaces=new Dictionary<string,List<HeightPatch>>();

            // Intersect with a*x+b*y+c >= 0. Clip a convex bounding rectangle first so
            // concave contours and holes are subsequently handled by the planar boolean engine.
            internal static PlanarRegion PositiveHeightRegion(PlanarRegion region,double a,double b,double c)
            {
                if(region.IsEmpty)return region;
                if(Math.Abs(a)+Math.Abs(b)<1e-14)return c>=0?region:PlanarRegion.Empty;
                var points=region.Paths.SelectMany(p=>p).ToList();
                double x0=points.Min(p=>p.X)/PlanarRegion.Scale,x1=points.Max(p=>p.X)/PlanarRegion.Scale;
                double y0=points.Min(p=>p.Y)/PlanarRegion.Scale,y1=points.Max(p=>p.Y)/PlanarRegion.Scale;
                var box=new[]{new[]{x0,y0},new[]{x1,y0},new[]{x1,y1},new[]{x0,y1}};
                var clipped=new List<double[]>();
                for(int i=0;i<4;i++)
                {
                    var p=box[i];var q=box[(i+1)%4];double u=a*p[0]+b*p[1]+c,v=a*q[0]+b*q[1]+c;
                    if(u>=0)clipped.Add(p);
                    if((u>=0)!=(v>=0)){double t=u/(u-v);clipped.Add(new[]{p[0]+t*(q[0]-p[0]),p[1]+t*(q[1]-p[1])});}
                }
                return clipped.Count<3?PlanarRegion.Empty:RegionIntersection(region,PlanarRegion.FromRings(new[]{clipped}));
            }

            internal static List<HeightPatch> LowestCeiling(PlanarRegion floor,double elevation,List<HeightPatch> candidates)
            {
                var faces=candidates.Select(p=>new HeightPatch{Region=PositiveHeightRegion(RegionIntersection(floor,p.Region),p.X,p.Y,p.C-elevation-1e-8),X=p.X,Y=p.Y,C=p.C,Top=true,RoofCeiling=p.RoofCeiling}).Where(p=>!p.Region.IsEmpty).ToList();
                var result=new List<HeightPatch>();
                for(int i=0;i<faces.Count;i++)
                {
                    var face=faces[i];var region=face.Region;
                    for(int j=0;j<faces.Count&&!region.IsEmpty;j++)
                    {
                        if(i==j)continue;var other=faces[j];
                        double a=face.X-other.X,b=face.Y-other.Y,c=face.C-other.C;
                        if(Math.Abs(a)+Math.Abs(b)<1e-12&&Math.Abs(c)<1e-7)
                        {if(j<i)region=RegionDifference(region,other.Region);continue;}
                        var lower=PositiveHeightRegion(other.Region,a,b,c);
                        region=RegionDifference(region,lower);
                    }
                    if(!region.IsEmpty)result.Add(new HeightPatch{Region=region,X=face.X,Y=face.Y,C=face.C,Top=true,RoofCeiling=face.RoofCeiling});
                }
                var cover=result.Select(p=>p.Region).Aggregate(PlanarRegion.Empty,RegionUnion);
                if(RegionDifference(floor,cover).Area*.09290304>.001)throw new InvalidOperationException("Над частью площади не найдена физическая нижняя поверхность перекрытия, потолка или кровли. Верхняя граница Room не использована вместо потолка.");
                return result;
            }

            private ClearanceGeometry RoomClearance(Record record,bool centre=false)
            {
                ClearanceGeometry review;string error;string cacheKey=record.Key+(centre?"/centre":"/net");
                if(clearanceErrors.TryGetValue(cacheKey,out error))throw new InvalidOperationException(error);
                if(!clearanceGeometry.TryGetValue(cacheKey,out review))
                {
                    try
                    {
                        List<HeightPatch> patches;
                        if(record.Element is Room)patches=PhysicalRoomClearance(record,centre);
                        else
                        {
                            if(UsesFamilyParameterArea(record.Element))throw new InvalidOperationException("Параметр площади семейства не задаёт высоту и уклон. Для высотной проверки требуется объёмная геометрия.");
                            patches=Solids(record.Element).SelectMany(s=>HeightPatches(s,record.Source.Transform)).ToList();
                            foreach(var patch in patches)patch.RoofCeiling=patch.Top;
                        }
                        var range=HeightRange(patches);var angles=patches.Where(p=>p.Top).Select(CeilingAngle).ToList();
                        review=new ClearanceGeometry{Patches=patches,Minimum=range[0],Maximum=range[1],MinimumSlope=angles.Min(),MaximumSlope=angles.Max()};
                        clearanceGeometry[cacheKey]=review;
                    }
                    catch(OperationCanceledException){throw;}
                    catch(Exception ex){error="Высота и уклон по геометрии: "+ex.Message;clearanceErrors[cacheKey]=error;throw new InvalidOperationException(error,ex);}
                }
                Notice("HEIGHT_GEOMETRY","Информация","Высота в свету: "+review.Minimum.ToString("0.###")+" - "+review.Maximum.ToString("0.###")+" м; уклон потолка: "+review.MinimumSlope.ToString("0.###")+" - "+review.MaximumSlope.ToString("0.###")+"°. По фактическим поверхностям; усреднение не применяется.",record);
                return review;
            }

            internal static double CeilingAngle(HeightPatch patch)
            {return Math.Atan(Math.Sqrt(patch.X*patch.X+patch.Y*patch.Y))*180/Math.PI;}

            private List<HeightPatch> PhysicalRoomClearance(Record record,bool centre=false)
            {
                var region=centre?RoomCentreRegion(record):RoomNetRegion(record);double z=RoomFloorElevation(record);
                var candidates=new List<HeightPatch>();var floorCover=PlanarRegion.Empty;
                var unsupported=new List<Tuple<PlanarRegion,double,string>>();
                foreach(var source in Sources.Where(s=>s.Loaded&&s.LoadError==null&&s.Mode!="exclude"))
                    foreach(var element in source.Elements.Where(e=>e is Floor||e is RoofBase||e is Ceiling||
                        e is Stairs||e is StairsRun||e.Category!=null&&IDHelper.ElIdValue(e.Category.Id)==(long)BuiltInCategory.OST_StructuralFraming||
                        e is FamilyInstance&&StairSetting((FamilyInstance)e)!=null))
                    {
                        var box=ElementBounds(element,source.Transform);
                        if(box==null||box[5]<z-1e-6||RegionIntersection(BoundsRegion(box),region).IsEmpty||!PhaseAccepted(source,element))continue;
                        string key=source.Key+"/"+element.UniqueId;List<HeightPatch> faces;
                        try
                        {
                            if(!clearanceFaces.TryGetValue(key,out faces))
                            {
                                faces=new List<HeightPatch>();
                                foreach(var solid in Solids(element))foreach(Face face in solid.Faces)
                                {
                                    var plane=face as PlanarFace;
                                    if(plane==null)
                                    {
                                        var cylinder=face as CylindricalFace;
                                        if(cylinder!=null&&Math.Abs(Math.Abs(source.Transform.OfVector(cylinder.Axis).Normalize().Z)-1)<1e-12)continue;
                                        throw new InvalidOperationException("Криволинейная поверхность требует отдельного точного обмера.");
                                    }
                                    var patch=HeightFace(plane,source.Transform);if(patch!=null){patch.RoofCeiling=element is RoofBase||element is Ceiling;faces.Add(patch);}
                                }
                                clearanceFaces[key]=faces;
                            }
                            foreach(var face in faces)
                            {
                                if(!face.Top)candidates.Add(face);
                                else if((element is Floor||element is RoofBase)&&Math.Abs(face.X)+Math.Abs(face.Y)<1e-12&&Math.Abs(face.C-z)<1e-6)
                                    floorCover=RegionUnion(floorCover,face.Region);
                            }
                        }
                        catch(OperationCanceledException){throw;}
                        catch(Exception ex){unsupported.Add(Tuple.Create(RegionIntersection(region,BoundsRegion(box)),box[2],source.Name+" / ID "+IDHelper.ElIdValue(element.Id)+": "+ex.Message));}
                    }
                if(RegionDifference(region,floorCover).Area*.09290304>.001)
                    throw new InvalidOperationException("Отметка низа помещения не подтверждена верхом пола на всей площади. Высота Room не подставлена вместо высоты в свету.");
                var top=LowestCeiling(region,z,candidates);
                foreach(var unknown in unsupported)
                {
                    var shield=top.Select(p=>PositiveHeightRegion(p.Region,-p.X,-p.Y,unknown.Item2-p.C-1e-6)).Aggregate(PlanarRegion.Empty,RegionUnion);
                    if(RegionDifference(unknown.Item1,shield).Area*.09290304>.001)throw new InvalidOperationException(unknown.Item3);
                }
                top.Add(new HeightPatch{Region=region,C=z,Top=false});return top;
            }

            private static bool NeedsCeilingCheck(Record record,Indicator metric)
            {return (RoomAreaMetric(metric)||GrossMetric(metric))&&OneOf(record.Role,"heated","auxiliary","public","niche","arch","mezzanine","residential-common","nonresidential-common");}

            private bool HasSlopedCeiling(Record record)
            {return RoomClearance(record).Patches.Any(p=>p.Top&&p.RoofCeiling&&CeilingAngle(p)>1e-7);}

            private PlanarRegion ClearanceMask(Record record,Indicator metric,double minimum,bool checkSlope,bool centre=false)
            {
                var review=RoomClearance(record,centre);bool slope=checkSlope&&review.Patches.Any(p=>p.Top&&p.RoofCeiling&&CeilingAngle(p)>1e-7);
                if(slope&&!IsPublic(record)&&record.Role!="public"&&record.Part!="nonresidential")
                    throw new InvalidOperationException("Обнаружен наклонный потолок. Для жилой мансарды в исходном ТЗ не согласованы нормативные пороги высоты; полный контур автоматически не засчитывается.");
                var result=PlanarRegion.Empty;
                foreach(var top in review.Patches.Where(p=>p.Top))foreach(var bottom in review.Patches.Where(p=>!p.Top))
                {
                    double required=minimum;
                    if(slope)required=Math.Max(required,metric==Indicator.PublicCalculated?1.5:PublicMansardHeight(CeilingAngle(top)));
                    var area=RegionIntersection(top.Region,bottom.Region);
                    result=RegionUnion(result,PositiveHeightRegion(area,top.X-bottom.X,top.Y-bottom.Y,top.C-bottom.C-required/.3048));
                }
                return result;
            }
        }
    }
}
