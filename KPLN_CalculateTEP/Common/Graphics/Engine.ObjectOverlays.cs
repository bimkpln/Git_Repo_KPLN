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
            public List<Issue> OverlayIssues {get;private set;} = new List<Issue>();
            private readonly Dictionary<string,PlanarRegion> editedObjectOverlays=new Dictionary<string,PlanarRegion>();
            private string ObjectOverlayKey(Record record){return Config.Method+"/"+(Config.Phase??"")+"/object/"+record.Key;}
            private static Color OverlayColour(string role)
            {
                if(OneOf(role,"heated","auxiliary","residential-common"))return new Color(76,175,110);
                if(OneOf(role,"loggia","balcony","terrace"))return new Color(56,162,214);
                if(role=="public")return new Color(233,174,59);
                if(role=="parking")return new Color(151,105,199);
                if(role=="technical-room")return new Color(124,140,153);
                return role==null||role=="unknown"?new Color(224,79,73):new Color(72,153,163);
            }
            internal static bool OverlayBoundaryChanged(PlanarRegion initial,PlanarRegion edited)
            {
                return RegionDifference(initial,edited).Area+RegionDifference(edited,initial).Area>1e-8;
            }
            private bool TryEditedObjectOverlay(Record record,out PlanarRegion result)
            {
                result=null;if(!Config.UseReviewRegions)return false;
                string key=ObjectOverlayKey(record);
                if(editedObjectOverlays.TryGetValue(key,out result))return result!=null;
                var binding=reviewBindings.SingleOrDefault(b=>b.Key==key);
                if(binding==null){editedObjectOverlays[key]=null;return false;}
                if(binding.Placement!=ReviewPlacement())throw new InvalidOperationException("Изменилось размещение источников цветовых областей. Перепривяжите проверенную область.");
                var region=doc.GetElement(binding.RegionId) as FilledRegion;
                if(region==null)throw new InvalidOperationException("Удалена расчётная цветовая область объекта. Привяжите новую область к этому объекту.");
                var view=doc.GetElement(region.OwnerViewId) as ViewPlan;
                if(view?.GenLevel==null||Math.Abs(view.GenLevel.ProjectElevation-record.Z)>FloorGeometryTolerance)
                    throw new InvalidOperationException("Цветовая область перенесена на другой расчётный этаж.");
                var actual=RegionFromLoops(region.GetBoundaries(),true);
                if(binding.InitialBoundary!=null&&!OverlayBoundaryChanged(PlanarRegion.FromRings(binding.InitialBoundary),actual))
                {editedObjectOverlays[key]=null;return false;}
                reviewInputs[key]=new ReviewInput{Key=key,Kind="Исправленная область объекта",Source=record.Source.Name,SourceKey=record.Source.Key,Apartment=record.Apartment,ObjectName=record.Element.Name,Building=record.Building,Section=record.Section,Element=IDHelper.ElIdValue(record.Element.Id).ToString(),Level=record.Level?.Name,
                    Elevation=record.Z,PlanElevation=record.Z,ObjectOverlay=true,Role=record.Role,Region=actual,Status="Исправленная область ID "+IDHelper.ElIdValue(region.Id)};
                editedObjectOverlays[key]=result=actual;
                Notice("EDITED_OBJECT_AREA","Предупреждение","Площадной контур объекта заменён изменённой цветовой областью ID "+IDHelper.ElIdValue(region.Id)+". Площадь, положение и пересечения читаются из её текущей границы; назначение и нормативные проверки сохранены.",record);
                return true;
            }
            private void FinalizeConfirmedReviewBindings()
            {
                var candidates=reviewBindings.Where(b=>b.RequiresReview&&reviewInputs.TryGetValue(b.Key,out var input)&&
                    !input.IsDraft&&input.Region!=null&&!input.Region.IsEmpty).ToList();
                if(candidates.Count==0)return;
                var previous=candidates.Select(b=>new {Binding=b,b.RequiresReview,b.DraftReason,b.DraftDescription}).ToList();
                var report=current??Last??new Run();bool committed=false;
                try
                {
                    var confirmed=new List<Tuple<ReviewBinding,ReviewInput,FilledRegion,View>>();
                    foreach(var binding in candidates)
                    {
                        var region=doc.GetElement(binding.RegionId) as FilledRegion;
                        if(region==null)continue;
                        var input=reviewInputs[binding.Key];
                        // Confirm the persisted outline, including when edited inputs were disabled for this run.
                        ValidateCreatedArea(input.Region,RegionFromLoops(region.GetBoundaries(),true));
                        var view=doc.GetElement(region.OwnerViewId) as View;
                        if(view==null)throw new InvalidOperationException("Не найден вид подтверждённой расчётной области ID "+IDHelper.ElIdValue(region.Id)+".");
                        confirmed.Add(Tuple.Create(binding,input,region,view));
                    }
                    if(confirmed.Count==0)return;
                    Transaction("ТЭП: подтверждение расчётных областей",report,()=>
                    {
                        var fill=new FilteredElementCollector(doc).OfClass(typeof(FillPatternElement)).Cast<FillPatternElement>().FirstOrDefault(f=>f.GetFillPattern().IsSolidFill);
                        foreach(var item in confirmed)
                        {
                            var binding=item.Item1;var input=item.Item2;var region=item.Item3;var view=item.Item4;
                            var colour=OverlayColour(input.Role);
                            var style=new OverrideGraphicSettings().SetProjectionLineColor(colour).SetSurfaceTransparency(65);
                            if(fill!=null)style.SetSurfaceForegroundPatternId(fill.Id).SetSurfaceForegroundPatternColor(colour);
                            if(input.Kind!=null&&input.Kind.StartsWith("Контур этажа"))style.SetSurfaceForegroundPatternVisible(false).SetSurfaceBackgroundPatternVisible(false);
                            view.SetElementOverrides(region.Id,style);
                            region.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)?.Set(input.Kind+" | "+input.Source+" | ID "+input.Element);
                            binding.RequiresReview=false;binding.DraftReason=null;binding.DraftDescription=null;
                        }
                        SaveReviewBindings();
                    });
                    committed=true;
                }
                catch(System.OperationCanceledException){throw;}
                catch(Autodesk.Revit.Exceptions.OperationCanceledException){throw;}
                catch(Autodesk.Revit.Exceptions.RegenerationFailedException){throw;}
                catch(Exception ex)
                {
                    report.Issue("EDITABLE_CONFIRM_STYLE","Предупреждение","Расчёт выполнен, но не удалось обновить оформление подтверждённых областей. Признак предварительного контура сохранён; при следующем расчёте области будут проверены повторно. "+ex.Message);
                }
                finally
                {
                    if(!committed)foreach(var item in previous)
                    {item.Binding.RequiresReview=item.RequiresReview;item.Binding.DraftReason=item.DraftReason;item.Binding.DraftDescription=item.DraftDescription;}
                }
            }
        }
    }
}
