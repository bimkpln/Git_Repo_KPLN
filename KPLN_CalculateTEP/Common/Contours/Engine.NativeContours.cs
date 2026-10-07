using Autodesk.Revit.DB;
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
            private readonly Dictionary<string,List<Tuple<int,Solid>>> nativeWallLayers=new Dictionary<string,List<Tuple<int,Solid>>>();
            private sealed class NativeShell
            {
                internal Record Record;
                internal List<Face> Exterior,Interior;
                internal PlanarRegion Section=PlanarRegion.Empty,Finish=PlanarRegion.Empty;
                internal double MinX,MinY,MaxX,MaxY;
            }
            private void ReadNativeWallLayers(List<Record> records)
            {
                var missing=records.Where(r=>!nativeWallLayers.ContainsKey(r.Key)).ToList();
                if(missing.Count==0)return;
                var pending=new Dictionary<string,List<Tuple<int,Solid>>>();
                var originals=missing.ToDictionary(r=>r.Key,r=>Solids(r.Element).Select(s=>SolidUtils.CreateTransformed(s,r.Source.Transform)).ToList());
                // Parts expose real layer geometry, including tapered and vertically compound walls.
                // The entire transaction is rolled back; no Parts or edits remain in the user's RVT.
                using(var transaction=new Transaction(doc,"ТЭП: временное чтение геометрии слоёв"))
                {
                    transaction.Start();
                    try
                    {
                        var identities=new Dictionary<string,LinkElementId>();
                        foreach(var r in missing)
                        {
                            LinkElementId identity;
                            if(r.Element.Document.Equals(doc))identity=new LinkElementId(r.Element.Id);
                            else
                            {
                                var link=r.Source.RootLink==null?null:doc.GetElement(r.Source.RootLink) as RevitLinkInstance;
                                if(link==null||!link.GetLinkDocument().Equals(r.Element.Document))throw new InvalidOperationException("Для чтения слоёв вложенной связи нужен доступ к непосредственному экземпляру связи: "+r.Source.Name);
                                identity=new LinkElementId(link.Id,r.Element.Id);
                            }
                            identities.Add(r.Key,identity);
                            if(PartUtils.GetAssociatedParts(doc,identity,true,true).Count==0)
                            {
                                if(!PartUtils.IsValidForCreateParts(doc,identity))throw new InvalidOperationException("Revit не предоставляет слои через Parts для стены ID "+IDHelper.ElIdValue(r.Element.Id));
                                Progress("Чтение слоёв стены: "+r.Source.Name+"; ID "+IDHelper.ElIdValue(r.Element.Id));
                                PartUtils.CreateParts(doc,new List<LinkElementId>{identity});
                            }
                        }
                        doc.Regenerate();
                        foreach(var r in missing)
                        {
                            var wall=(Wall)r.Element;var structure=wall.WallType.GetCompoundStructure();
                            if(structure==null||structure.GetFirstCoreLayerIndex()<0)throw new InvalidOperationException("Для стены ID "+IDHelper.ElIdValue(wall.Id)+" в модели не определено ядро.");
                            var parts=PartUtils.GetAssociatedParts(doc,identities[r.Key],false,true).Select(id=>doc.GetElement(id)).OfType<Part>().ToList();
                            var indexed=new List<Tuple<int,Part>>();
                            BuiltInParameter layerIndex;
                            if(!Enum.TryParse("DPART_LAYER_INDEX",out layerIndex))
                                throw new InvalidOperationException("Эта версия Revit не предоставляет индекс слоя Parts. Сложное ядро нельзя однозначно сопоставить слоям типа стены.");
                            foreach(var part in parts)
                            {
                                var p=part.get_Parameter(layerIndex);int index;
                                if(p==null||!(p.StorageType==StorageType.Integer?ParseLayerIndex(p.AsInteger().ToString(CultureInfo.InvariantCulture),out index):ParseLayerIndex(p.AsString(),out index)))
                                    throw new InvalidOperationException("Revit не вернул однозначный индекс слоя Part, стена ID "+IDHelper.ElIdValue(wall.Id));
                                indexed.Add(Tuple.Create(index,part));
                            }
                            var expected=structure.GetLayers().Select((layer,i)=>new{layer,i}).Where(x=>x.layer.Width>0).Select(x=>x.i).ToList();
                            var present=indexed.Select(x=>x.Item1).Distinct().OrderBy(x=>x).ToList();
                            var offsets=new[]{0,1}.Where(offset=>present.SequenceEqual(expected.Select(i=>i+offset))).ToList();
                            if(offsets.Count!=1)throw new InvalidOperationException("Состав Parts не соответствует слоям стены ID "+IDHelper.ElIdValue(wall.Id)+"; индексы слоёв не угадываются.");
                            var geometry=new List<Tuple<int,Solid>>();
                            foreach(var item in indexed)
                                foreach(var solid in Solids(item.Item2))geometry.Add(Tuple.Create(item.Item1-offsets[0],SolidUtils.Clone(solid)));
                            if(geometry.Count==0)throw new InvalidOperationException("Нет геометрии слоёв стены ID "+IDHelper.ElIdValue(wall.Id));
                            var original=originals[r.Key];
                            double originalVolume=original.Sum(s=>s.Volume),partsVolume=geometry.Sum(x=>x.Item2.Volume);
                            if(originalVolume<=0||Math.Abs(originalVolume-partsVolume)>Math.Max(1e-7,originalVolume*1e-7))
                                throw new InvalidOperationException("Parts не сохранили объём исходной стены ID "+IDHelper.ElIdValue(wall.Id)+". Изменённые части не приняты за исходные слои.");
                            var originalCenter=original.Aggregate(XYZ.Zero,(a,s)=>a+s.ComputeCentroid()*s.Volume)/originalVolume;
                            var partsCenter=geometry.Aggregate(XYZ.Zero,(a,x)=>a+x.Item2.ComputeCentroid()*x.Item2.Volume)/partsVolume;
                            if(originalCenter.DistanceTo(partsCenter)>1e-5)throw new InvalidOperationException("Размещение Parts не соответствует экземпляру исходной стены / связи.");
                            pending.Add(r.Key,geometry);
                        }
                    }
                    finally
                    {
                        if(transaction.GetStatus()==TransactionStatus.Started&&transaction.RollBack()!=TransactionStatus.RolledBack)
                            throw new InvalidOperationException("Revit не подтвердил откат временного чтения слоёв. Расчёт остановлен.");
                        DisposeSpatialCalculators();localSpatialVolumes.Clear();floorWallFaces.Clear();floorFaceHeights.Clear();
                    }
                }
                foreach(var pair in pending)nativeWallLayers[pair.Key]=pair.Value;
            }
            private static bool ParseLayerIndex(string value,out int index)
            {return int.TryParse(value,NumberStyles.Integer,CultureInfo.InvariantCulture,out index);}

            private static bool OnShell(NativeShell shell,List<Face> faces,XYZ point)
            {
                double tolerance=ArcChordTolerance*2+2/PlanarRegion.Scale;
                if(point.X<shell.MinX-tolerance||point.X>shell.MaxX+tolerance||point.Y<shell.MinY-tolerance||point.Y>shell.MaxY+tolerance)return false;
                var local=shell.Record.Source.Transform.Inverse.OfPoint(point);
                foreach(var face in faces)
                {
                    var projection=face.Project(local);
                    if(projection!=null&&projection.Distance<=tolerance&&face.IsInside(projection.UVPoint))return true;
                }
                return false;
            }
            private static bool OnPlanBoundary(PlanarRegion region,XYZ point)
            {
                if(region==null||region.IsEmpty)return false;
                var p=new[]{point.X,point.Y,0.0};
                foreach(var ring in region.Paths)
                    for(int i=0;i<ring.Count;i++)
                    {
                        var a=ring[i];var b=ring[(i+1)%ring.Count];
                        if(RationalContours.DistanceToSegment(p,new[]{a.X/PlanarRegion.Scale,a.Y/PlanarRegion.Scale,0.0},new[]{b.X/PlanarRegion.Scale,b.Y/PlanarRegion.Scale,0.0})<=ArcChordTolerance*2+2/PlanarRegion.Scale)return true;
                    }
                return false;
            }
            private PlanarRegion NativeWallContourRegion(List<Record> records,string boundary,double elevation)
            {
                if(boundary=="core")ReadNativeWallLayers(records);
                var shells=new List<NativeShell>();var all=PlanarRegion.Empty;
                foreach(var r in records)
                {
                    Progress("Сечения ограждений: "+r.Source.Name+"; ID "+IDHelper.ElIdValue(r.Element.Id));
                    var wall=(Wall)r.Element;
                    var shell=new NativeShell{Record=r,
                        Exterior=HostObjectUtils.GetSideFaces(wall,ShellLayerType.Exterior).Select(wall.GetGeometryObjectFromReference).OfType<Face>().ToList(),
                        Interior=HostObjectUtils.GetSideFaces(wall,ShellLayerType.Interior).Select(wall.GetGeometryObjectFromReference).OfType<Face>().ToList()};
                    if(shell.Exterior.Count==0||shell.Interior.Count==0)throw new InvalidOperationException("Revit не определил наружную / внутреннюю сторону стены ID "+IDHelper.ElIdValue(wall.Id));
                    if(boundary=="core")
                    {
                        int first=wall.WallType.GetCompoundStructure().GetFirstCoreLayerIndex();
                        foreach(var layer in nativeWallLayers[r.Key])
                        {
                            // Parts created in the host already contain the link instance transform.
                            var plan=SectionRegion(layer.Item2,elevation);
                            if(layer.Item1<first)shell.Finish=RegionUnion(shell.Finish,plan);
                            else shell.Section=RegionUnion(shell.Section,plan);
                        }
                    }
                    else foreach(var solid in Solids(wall))shell.Section=RegionUnion(shell.Section,SectionRegion(SolidUtils.CreateTransformed(solid,r.Source.Transform),elevation));
                    if(shell.Section.IsEmpty)throw new InvalidOperationException("Пустое сечение ограждения на отметке обмера: ID "+IDHelper.ElIdValue(wall.Id));
                    var points=shell.Section.Paths.SelectMany(x=>x).ToList();
                    shell.MinX=points.Min(p=>p.X)/PlanarRegion.Scale;shell.MaxX=points.Max(p=>p.X)/PlanarRegion.Scale;
                    shell.MinY=points.Min(p=>p.Y)/PlanarRegion.Scale;shell.MaxY=points.Max(p=>p.Y)/PlanarRegion.Scale;
                    all=RegionUnion(all,shell.Section);shells.Add(shell);
                }
                if(all.IsEmpty)throw new InvalidOperationException("Не получены сечения наружных ограждений.");
                var rings=new List<List<double[]>>();
                foreach(var ring in all.Paths)
                {
                    bool outside=false,inside=false,unclassified=false;
                    for(int i=0;i<ring.Count;i++)
                    {
                        planarCheckpoint?.Invoke();var a=ring[i];var b=ring[(i+1)%ring.Count];
                        var p=new XYZ(((double)a.X+b.X)/2/PlanarRegion.Scale,((double)a.Y+b.Y)/2/PlanarRegion.Scale,elevation);
                        bool outer=shells.Any(s=>boundary=="core"&&!s.Finish.IsEmpty?OnPlanBoundary(s.Finish,p):OnShell(s,s.Exterior,p));
                        bool inner=shells.Any(s=>OnShell(s,s.Interior,p));
                        outside|=outer;inside|=inner;unclassified|=!outer&&!inner;
                    }
                    if(outside&&inside||unclassified)throw new InvalidOperationException("Сечение ограждений имеет незамкнутую оболочку, проём либо участок без однозначной стороны. Контур по осям не подставлен; требуется восстановление границы проёма.");
                    if(boundary=="interior"?inside:outside)rings.Add(ring.Select(p=>new[]{p.X/PlanarRegion.Scale,p.Y/PlanarRegion.Scale}).ToList());
                }
                if(rings.Count==0)throw new InvalidOperationException("Сечение ограждений не содержит замкнутой нормативной границы.");
                return PlanarRegion.FromRings(rings);
            }
            private Solid WallContour(List<Record> records,Metric metric,Indicator indicator,double? floorElevation=null)
            { return BuildPlanarSolid(WallContourRegion(records, metric, indicator, floorElevation), 0, 1, true); }
            private PlanarRegion WallContourRegion(List<Record> records,Metric metric,Indicator indicator,double? floorElevation=null)
            {
                double elevation=floorElevation??records.First().Z+RequiredNumber(metric.WallCutHeight,"высота сечения стен, м")/.3048;
                string boundary=metric.WallBoundary;
                if(boundary=="auto")boundary=(int)indicator<=4?Methodology.GnsWallBoundary:GrossMetric(indicator)?"interior":"exterior";
                var failures=new List<string>();
                // Cheap exact analytic path remains useful for ordinary walls with door openings.
                // Native sections remove its restrictions on axes, surface types and layer layout.
                try{return AnalyticWallContourRegion(records,metric,indicator,elevation);}
                catch(OperationCanceledException){throw;}
                catch(Autodesk.Revit.Exceptions.OperationCanceledException){throw;}
                catch(Autodesk.Revit.Exceptions.RegenerationFailedException){throw;}
                catch(Exception ex){failures.Add("Аналитические поверхности: "+ex.Message);}
                try
                {
                    var result=NativeWallContourRegion(records,boundary,elevation);
                    Notice("NATIVE_ENVELOPE","Информация","Контур этажа получен из сечений фактической геометрии ограждений; замкнутость осей не требуется.",metric:indicator.ToString());
                    return result;
                }
                catch(OperationCanceledException){throw;}
                catch(Autodesk.Revit.Exceptions.OperationCanceledException){throw;}
                catch(Autodesk.Revit.Exceptions.RegenerationFailedException){throw;}
                catch(Exception ex){failures.Add("Геометрические сечения: "+ex.Message);}
                throw new InvalidOperationException(string.Join(" | ",failures));
            }
        }
    }
}
