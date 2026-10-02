using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_CalculateTEP.Common;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using PC = KPLN_CalculateTEP.Common.Geometry.TepClipper.Clipper;
using CP = KPLN_CalculateTEP.Common.Geometry.TepClipper.IntPoint;
using CT = KPLN_CalculateTEP.Common.Geometry.TepClipper.ClipType;
using PT = KPLN_CalculateTEP.Common.Geometry.TepClipper.PolyType;
using PF = KPLN_CalculateTEP.Common.Geometry.TepClipper.PolyFillType;

using TepClipper = KPLN_CalculateTEP.Common.Geometry.TepClipper;
using KPLN_CalculateTEP.Common.Methodologies;
namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public partial class Engine
        {
            private sealed class VolumeRetryBudget
            {
                private int remaining=64;
                internal void Spend()
                {if(remaining--<=0)throw new InvalidOperationException("Исчерпан лимит 64 резервных геометрических операций для одной пары тел.");}
            }
            private sealed class VolumeFragment
            {
                internal string RecordKey;
                internal Solid Shape,Above,Below;
                internal bool Split;
            }
            private sealed class VolumeSet
            {
                internal readonly List<VolumeFragment> Fragments=new List<VolumeFragment>();
                internal readonly List<Tuple<string,string,string>> Problems=new List<Tuple<string,string,string>>();
                internal bool Reconstructed;
            }
            private void DisposeSpatialCalculators()
            {
                foreach(var calculator in spatialCalculators.Values)calculator.Dispose();
                spatialCalculators.Clear();
            }
            private void ClearVolumeCaches()
            {
                DisposeSpatialCalculators();localSpatialVolumes.Clear();rawVolumeSolids.Clear();worldVolumeSolids.Clear();volumeSets.Clear();
            }
            private static void SignaturePart(StringBuilder builder,string value)
            {if(value==null)builder.Append("-1:");else builder.Append(value.Length).Append(':').Append(value);}
            private string VolumeInputKey(List<Record> records)
            {
                // Per-run cache. Source transforms and document geometry cannot change during the read phase.
                // Include resolved rule/level state, and rebind cached fragments to this metric's records.
                var key=new StringBuilder();SignaturePart(key,Config.Method);SignaturePart(key,Config.Phase);
                foreach(var r in records.OrderBy(x=>x.Key,StringComparer.Ordinal))
                    foreach(var value in new[]{r.Key,r.Source.Mode,r.Building,r.Section,r.Profile,r.BuildingClass,r.Role,r.Part,r.Apartment,r.Vertical,
                        r.Override.HasValue?r.Override.Value.ToString():null,r.Factor?.ToString("R",CultureInfo.InvariantCulture),r.Manual.ToString(),r.SingleStorey.ToString(),r.BuildingIncluded.ToString(),
                        r.Level?.Key,r.Level?.Include.ToString(),r.Level?.Kind,r.Level?.Above,r.Level?.TopSlab,r.Level?.Height,r.Level?.RoofRatio,r.Level?.RoofArea,
                        r.Z.ToString("R",CultureInfo.InvariantCulture)})SignaturePart(key,value);
                return key.ToString();
            }
            private List<Solid> VolumeSolids(Record r,string metric)
            {
                List<Solid> world;if(worldVolumeSolids.TryGetValue(r.Key,out world))return world;
                List<Solid> local;
                if(!rawVolumeSolids.TryGetValue(r.Element,out local))
                {
                    Progress("Геометрия конструкций: "+r.Source.Name+"; ID "+IDHelper.ElIdValue(r.Element.Id));
                    local=Measure("Геометрия конструкций",r,metric,()=>Solids(r.Element));rawVolumeSolids[r.Element]=local;
                }
                world=local.Select(s=>SolidUtils.CreateTransformed(s,r.Source.Transform)).ToList();worldVolumeSolids[r.Key]=world;return world;
            }
            private Solid RemoveNearby(Solid shape,PlanIndex index,Record r,string metric,string stage)
            {
                foreach(var other in index.Query(shape))
                {
                    Progress(stage+": "+r.Building+" / "+r.Source.Name+"; ID "+IDHelper.ElIdValue(r.Element.Id));
                    try{shape=Measure(stage,r,metric,()=>Subtract(shape,other));}
                    catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){throw new InvalidOperationException("Этап: "+stage+". Объект A: ["+BooleanObject(r)+"]. Объект B: ["+BooleanObject(index.OwnerOf(other))+"]. "+ex.Message,ex);}
                    if(shape==null||shape.Volume<1e-9)break;
                }
                return shape;
            }
            private static double CheckedVolume(Solid solid)
            {
                if(solid==null)return 0;
                double value=solid.Volume;
                if(double.IsNaN(value)||double.IsInfinity(value)||value<0)throw new InvalidOperationException("Недопустимый объём результата геометрической операции: "+value.ToString("R",CultureInfo.InvariantCulture)+" фут³.");
                return value;
            }
            private static double VolumeTolerance(double value){return Math.Max(1e-7,Math.Abs(value)*1e-8);}
            private static List<Solid> DifferenceResult(Solid original,Solid difference,Solid intersection=null)
            {
                if(difference==null)throw new InvalidOperationException("Revit не вернул тело результата вычитания.");
                double before=CheckedVolume(original),after=CheckedVolume(difference),tolerance=VolumeTolerance(before);
                if(after>before+tolerance)throw new InvalidOperationException("Вычитание увеличило объём исходного тела.");
                if(intersection!=null&&Math.Abs(before-CheckedVolume(intersection)-after)>tolerance)
                    throw new InvalidOperationException("Нарушен баланс объёма после вычитания пересечения.");
                return after>0?new List<Solid>{difference}:new List<Solid>();
            }
            private Solid VolumeBoolean(Solid a,Solid b,BooleanOperationsType operation,Record r,string metric,VolumeRetryBudget budget)
            {
                budget?.Spend();
                string stage=operation==BooleanOperationsType.Intersect?"Проверка фактического пересечения объёмов":"Вычитание объёмов";
                Progress(stage+": "+r.Building+" / "+r.Source.Name+"; ID "+IDHelper.ElIdValue(r.Element.Id));
                return Measure(stage,r,metric,()=>BooleanOperationsUtils.ExecuteBooleanOperation(a,b,operation));
            }
            private List<Solid> ConnectedVolumeParts(Solid solid,Record r,string metric,VolumeRetryBudget budget)
            {
                budget.Spend();Progress("Разбиение объёма на связные части: "+r.Building+"; ID "+IDHelper.ElIdValue(r.Element.Id));
                var parts=Measure("Разбиение объёмов на связные части",r,metric,()=>SolidUtils.SplitVolumes(solid).ToList());
                double before=CheckedVolume(solid),after=parts.Sum(CheckedVolume);
                if(Math.Abs(before-after)>VolumeTolerance(before)||before>0&&!parts.Any(p=>CheckedVolume(p)>0))
                    throw new InvalidOperationException("Разбиение на связные части не сохранило объём исходного тела.");
                return parts.Where(p=>CheckedVolume(p)>0).ToList();
            }
            private List<Solid> NormalizedVolumeDifference(Solid a,Solid b,Record r,string metric,VolumeRetryBudget budget)
            {
                var bounds=Bounds(a);var origin=new XYZ((bounds[0]+bounds[3])/2,(bounds[1]+bounds[4])/2,(bounds[2]+bounds[5])/2);
                var errors=new List<string>();
                foreach(double scale in new[]{1.0,1000.0})
                {
                    try
                    {
                        Progress("Вычитание в локальных координатах: "+r.Building+"; ID "+IDHelper.ElIdValue(r.Element.Id));
                        var transform=Transform.Identity;transform.BasisX=XYZ.BasisX*scale;transform.BasisY=XYZ.BasisY*scale;transform.BasisZ=XYZ.BasisZ*scale;transform.Origin=-origin*scale;
                        var left=SolidUtils.CreateTransformed(a,transform);var right=SolidUtils.CreateTransformed(b,transform);double cube=scale*scale*scale;
                        if(Math.Abs(CheckedVolume(left)/cube-CheckedVolume(a))>VolumeTolerance(CheckedVolume(a))||Math.Abs(CheckedVolume(right)/cube-CheckedVolume(b))>VolumeTolerance(CheckedVolume(b)))
                            throw new InvalidOperationException("Преобразование изменило исходный объём.");
                        var difference=VolumeBoolean(left,right,BooleanOperationsType.Difference,r,metric,budget);
                        var intersection=VolumeBoolean(left,right,BooleanOperationsType.Intersect,r,metric,budget);
                        if(intersection==null||CheckedVolume(intersection)>Math.Min(CheckedVolume(left),CheckedVolume(right))+VolumeTolerance(CheckedVolume(left)))throw new InvalidOperationException("Некорректное контрольное пересечение.");
                        var verified=DifferenceResult(left,difference,intersection);var restored=new List<Solid>();
                        foreach(var part in verified)restored.Add(SolidUtils.CreateTransformed(part,transform.Inverse));
                        if(Math.Abs(restored.Sum(CheckedVolume)+CheckedVolume(intersection)/cube-CheckedVolume(a))>VolumeTolerance(CheckedVolume(a)))
                            throw new InvalidOperationException("Обратное преобразование нарушило баланс объёмов.");
                        booleanNormalizedRecoveries++;return restored;
                    }
                    catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){errors.Add("Масштаб "+scale.ToString(CultureInfo.InvariantCulture)+": "+ex.Message);}
                }
                throw new InvalidOperationException(string.Join(" | ",errors));
            }
            private List<Solid> SubtractVolume(Solid a,Solid b,Record r,string metric,VolumeRetryBudget budget=null,bool allowSplit=true)
            {
                if(a==null||CheckedVolume(a)==0)return new List<Solid>();
                if(b==null||CheckedVolume(b)==0||!Overlaps(Bounds(a),Bounds(b)))return new List<Solid>{a};
                var failures=new List<string>();List<Solid> planarResult;string planarFailure;
                if(TryLayeredDifference(a,b,out planarResult,out planarFailure))return planarResult;
                failures.Add("Плоское разложение: "+planarFailure);
                try{return DifferenceResult(a,VolumeBoolean(a,b,BooleanOperationsType.Difference,r,metric,budget));}
                catch(System.OperationCanceledException){throw;}
                catch(Exception ex){failures.Add("Прямое вычитание: "+ex.Message);}
                budget=budget??new VolumeRetryBudget();
                Solid intersection=null;
                try
                {
                    intersection=VolumeBoolean(a,b,BooleanOperationsType.Intersect,r,metric,budget);
                    if(intersection==null)throw new InvalidOperationException("Revit не вернул результат проверки пересечения.");
                    double overlap=CheckedVolume(intersection);
                    // Only an exactly empty intersection skips subtraction. Never dismiss a small positive overlap.
                    if(overlap==0){booleanTouchSkips++;return new List<Solid>{a};}
                    if(overlap>Math.Min(CheckedVolume(a),CheckedVolume(b))+VolumeTolerance(CheckedVolume(a)))
                        throw new InvalidOperationException("Пересечение больше одного из исходных тел.");
                }
                catch(System.OperationCanceledException){throw;}
                catch(Exception ex){intersection=null;failures.Add("Проверка пересечения: "+ex.Message);}
                if(allowSplit)
                {
                    try
                    {
                        var left=ConnectedVolumeParts(a,r,metric,budget);var right=ConnectedVolumeParts(b,r,metric,budget);
                        if(left.Count>1||right.Count>1)
                        {
                            var recovered=new List<Solid>();
                            foreach(var part in left)
                            {
                                var remaining=new List<Solid>{part};
                                foreach(var cutter in right)
                                {
                                    var next=new List<Solid>();
                                    foreach(var piece in remaining)next.AddRange(SubtractVolume(piece,cutter,r,metric,budget,false));
                                    remaining=next;if(remaining.Count==0)break;
                                }
                                recovered.AddRange(remaining);
                            }
                            double volume=recovered.Sum(CheckedVolume),before=CheckedVolume(a);
                            if(volume>before+VolumeTolerance(before)||intersection!=null&&Math.Abs(before-CheckedVolume(intersection)-volume)>VolumeTolerance(before))
                                throw new InvalidOperationException("Нарушен баланс объёма после вычитания связных частей.");
                            booleanSplitRecoveries++;return recovered;
                        }
                    }
                    catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){failures.Add("Вычитание связных частей: "+ex.Message);}
                }
                if(intersection!=null)
                {
                    try
                    {
                        var recovered=DifferenceResult(a,VolumeBoolean(a,intersection,BooleanOperationsType.Difference,r,metric,budget),intersection);
                        booleanIntersectionRecoveries++;return recovered;
                    }
                    catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){failures.Add("Вычитание A ∩ B: "+ex.Message);}
                }
                try{return NormalizedVolumeDifference(a,b,r,metric,budget);}
                catch(System.OperationCanceledException){throw;}
                catch(Exception ex){failures.Add("Локальные координаты: "+ex.Message);}
                throw new InvalidOperationException("Не удалось точно вычесть объёмы после резервных попыток. "+string.Join(" | ",failures));
            }
            private List<Solid> RemoveNearbyVolumes(List<Solid> shapes,PlanIndex index,Record r,string metric,string stage="Вычитание объёмов")
            {
                var result=new List<Solid>();
                foreach(var original in shapes.Where(s=>s!=null&&CheckedVolume(s)>0))
                {
                    var remaining=new List<Solid>{original};
                    // All recovered pieces are subsets of original, so its candidate set remains sufficient.
                    foreach(var other in index.Query(original))
                    {
                        var next=new List<Solid>();
                        foreach(var piece in remaining)
                        {
                            try{next.AddRange(SubtractVolume(piece,other,r,metric));}
                            catch(System.OperationCanceledException){throw;}
                            catch(Exception ex){throw new InvalidOperationException("Этап: "+stage+". Объект A: ["+BooleanObject(r)+"]. Объект B: ["+BooleanObject(index.OwnerOf(other))+"]. "+ex.Message,ex);}
                        }
                        remaining=next;if(remaining.Count==0)break;
                    }
                    result.AddRange(remaining);
                }
                return result;
            }
            private VolumeSet BuildVolumeSet(List<Record> records,string metric)
            {
                var result=new VolumeSet();var inputs=new List<Tuple<Record,Solid>>();
                var ordered=records.OrderBy(r=>r.Key,StringComparer.Ordinal).ToList();
                var levels=records.Where(r=>r.Level!=null&&r.Level.Include&&r.Level.Kind!="exclude").Select(r=>r.Z).ToList();
                if(levels.Count==0)throw new InvalidOperationException("Не определён нижний расчётный этаж.");
                double lowest=levels.Min();
                var envelopes=ordered.Where(r=>r.Role=="envelope"&&r.Override!=false).ToList();
                result.Reconstructed=envelopes.Count==0;bool foundSpace=false;
                int processed=0;
                foreach(var r in envelopes.Count>0?envelopes:ordered)
                {
                    Progress("Подготовка объёмов: "+(++processed)+" / "+(envelopes.Count>0?envelopes.Count:ordered.Count)+"; ID "+IDHelper.ElIdValue(r.Element.Id));
                    if(r.Override==false)continue;
                    try
                    {
                        if(envelopes.Count>0)
                        {
                            var solids=VolumeSolids(r,metric);if(solids.Count==0)throw new InvalidOperationException("У расчётной оболочки нет замкнутого тела.");
                            foreach(var solid in solids)inputs.Add(Tuple.Create(r,solid));
                        }
                        else if(r.Element is SpatialElement&&!(r.Element is Area)&&!OneOf(r.Role,"balcony","terrace","canopy","porch","pit","passage","ventilated-void","soil-filled","decoration","french-balcony"))
                        {inputs.Add(Tuple.Create(r,SpatialVolume(r)));foundSpace=true;}
                        else if(r.Role=="structure")foreach(var solid in VolumeSolids(r,metric))inputs.Add(Tuple.Create(r,solid));
                    }
                    catch(System.OperationCanceledException){throw;}
                    catch(Exception ex)
                    {
                        // An incomplete explicit envelope cannot stand in for the intended closed building body.
                        if(envelopes.Count>0)throw new InvalidOperationException("Оболочка "+r.Source.Name+" / ID "+IDHelper.ElIdValue(r.Element.Id)+": "+ex.Message,ex);
                        result.Problems.Add(Tuple.Create(r.Key,"VOLUME_PART",ex.Message));Notice("VOLUME_PART","Ошибка",ex.Message,r,metric);
                    }
                }
                if(result.Reconstructed&&!foundSpace)throw new InvalidOperationException("Нет замкнутого внешнего объёма или объёмов Rooms/Spaces. Сумма материалов стен и перекрытий не является строительным объёмом.");
                if(inputs.Count==0)throw new InvalidOperationException("Расчётная оболочка пуста.");
                var masks=new PlanIndex(true);
                foreach(var r in ordered.Where(r=>r.Override==false||OneOf(r.Role,"balcony","terrace","canopy","passage","decoration","ventilated-void","soil-filled")))
                {
                    Progress("Подготовка исключений объёма: "+r.Building+"; ID "+IDHelper.ElIdValue(r.Element.Id));
                    try
                    {
                        if(r.Element is SpatialElement&&!(r.Element is Area))masks.Add(SpatialVolume(r),r);
                        else foreach(var solid in VolumeSolids(r,metric))masks.Add(solid,r);
                    }
                    catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){Notice("VOLUME_MASK","Ошибка",ex.Message,r,metric);throw new InvalidOperationException("Не удалось построить исключаемый объём "+r.Source.Name+" / ID "+IDHelper.ElIdValue(r.Element.Id)+": "+ex.Message,ex);}
                }
                var expandedInputs=new List<Tuple<Record,Solid>>();
                foreach(var input in inputs)foreach(var part in DecomposeVolumeInput(input.Item2,input.Item1,metric))expandedInputs.Add(Tuple.Create(input.Item1,part));
                inputs=expandedInputs;
                var used=new PlanIndex(true);processed=0;
                foreach(var input in inputs)
                {
                    var r=input.Item1;
                    Progress("Объёмные фрагменты: "+(++processed)+" / "+inputs.Count+"; ID "+IDHelper.ElIdValue(r.Element.Id));
                    try
                    {
                        // Clipping, differences and union distribution commute. Never construct a whole-building solid.
                        var shape=Measure("Обрезка по нижнему этажу",r,metric,()=>Half(input.Item2,lowest,true));
                        var fragments=RemoveNearbyVolumes(new List<Solid>{shape},masks,r,metric,"Исключаемые помещения / элементы");
                        fragments=RemoveNearbyVolumes(fragments,used,r,metric,"Устранение повторного учёта пересекающихся тел");
                        // Publish only after every subtraction for this input succeeds; never keep a partial retry.
                        foreach(var fragment in fragments.Where(f=>CheckedVolume(f)>=1e-9))
                        {used.Add(fragment,r);result.Fragments.Add(new VolumeFragment{RecordKey=r.Key,Shape=fragment});}
                    }
                    catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){result.Problems.Add(Tuple.Create(r.Key,"VOLUME_BOOLEAN",ex.Message));Notice("VOLUME_BOOLEAN","Ошибка",ex.Message,r,metric);}
                }
                return result;
            }
            private VolumeSet BuildingVolume(List<Record> records,string metric)
            {
                string key=VolumeInputKey(records);VolumeSet result;
                if(volumeSets.TryGetValue(key,out result))
                {volumeCacheHits++;Progress("Повторное использование объёмных фрагментов: "+records[0].Building+" / "+MetricLabel(metric));}
                else
                {
                    result=Measure("Подготовка объёмных фрагментов",null,metric,()=>BuildVolumeSet(records,metric));
                    volumeSets.Add(key,result);
                }
                var lookup=records.ToDictionary(r=>r.Key,StringComparer.Ordinal);
                if(result.Reconstructed)Notice("VOLUME_ENVELOPE","Предупреждение","Объём восстановлен из пространств и ограждений с однократным учётом пересечений. Проверьте полноту внешней оболочки, в том числе чердаки и надстройки; для подтверждённого результата задайте замкнутую расчётную оболочку.",records[0],metric);
                // Cached partial results must replay their diagnostics for every affected indicator.
                foreach(var problem in result.Problems)Notice(problem.Item2,"Ошибка",problem.Item3,lookup[problem.Item1],metric);
                return result;
            }
            private Solid VolumeSlice(VolumeFragment fragment,Record r,Indicator metric)
            {
                if(metric==Indicator.Volume)return fragment.Shape;
                if(!fragment.Split)
                {
                    Progress("Разделение по нулю здания: "+r.Building+"; ID "+IDHelper.ElIdValue(r.Element.Id));
                    Measure("Разделение надземного / подземного объёма",r,metric.ToString(),()=>
                    {
                        fragment.Above=Half(fragment.Shape,zero,true);fragment.Below=Half(fragment.Shape,zero,false);
                        double total=fragment.Shape.Volume,parts=(fragment.Above?.Volume??0)+(fragment.Below?.Volume??0);
                        if(Math.Abs(total-parts)>Math.Max(1e-6,total*1e-7))throw new InvalidOperationException("Нарушен баланс надземного и подземного фрагментов после разреза по нулю здания.");
                        fragment.Split=true;return true;
                    });
                }
                return metric==Indicator.VolumeAbove?fragment.Above:fragment.Below;
            }
            private void CalculateVolumes(List<Record> records,Indicator metric)
            {
                foreach(var group in records.Where(r=>r.Source.Mode!="exclude").GroupBy(r=>new{r.Building,r.Section,Reference=r.Source.Mode=="reference",ReferenceKey=r.Source.Mode=="reference"?r.Source.Key:""}))
                {
                    try
                    {
                        var list=group.ToList();var volume=BuildingVolume(list,metric.ToString());var lookup=list.ToDictionary(r=>r.Key,StringComparer.Ordinal);int processed=0;
                        foreach(var fragment in volume.Fragments)
                        {
                            var r=lookup[fragment.RecordKey];
                            Progress("Результат объёма: "+(++processed)+" / "+volume.Fragments.Count+"; ID "+IDHelper.ElIdValue(r.Element.Id));
                            try
                            {
                                var shape=VolumeSlice(fragment,r,metric);if(shape==null||shape.Volume<1e-9)continue;
                                current.Details.Add(Row(r,metric,shape.Volume*.028316846592,1,group.Key.Reference,group.Key.Reference?"Справочная геометрия источника, без включения в итог":"Фрагмент внешнего объёма исходного объекта; исключения вычтены, пересечения учтены один раз",shape,true));
                            }
                            catch(System.OperationCanceledException){throw;}
                            catch(Exception ex){Notice("VOLUME_SLICE","Ошибка",ex.Message,r,metric.ToString());}
                        }
                    }
                    catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){Notice("VOLUME_INPUT","Ошибка",ex.Message,group.First(),metric.ToString());}
                }
            }
        }
    }
}
