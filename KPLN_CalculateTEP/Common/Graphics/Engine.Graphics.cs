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
            private void AddErrorContours()
            {
                if(!ViewsEnabled)return;
                foreach(var issue in current.Issues.Where(x=>x.Severity=="Ошибка"&&!string.IsNullOrWhiteSpace(x.Element)).GroupBy(x=>x.Source+"/"+x.Element).Select(g=>g.First()).ToList())
                {
                    try
                    {
                        var source=Sources.FirstOrDefault(s=>s.Name==issue.Source&&s.Loaded);long id;
                        if(source==null||!long.TryParse(issue.Element,out id))continue;
                        var element=source.Document.GetElement(IDHelper.CreateElementId(id));if(element==null)continue;
                        var level=source.Document.GetElement(element.LevelId) as Level;
                        var r=new Record{Source=source,Element=element,Building=string.IsNullOrEmpty(issue.Building)?"Не определён":issue.Building,
                            Level=level==null?null:Config.Levels.FirstOrDefault(l=>l.Key==source.Key+"/"+level.UniqueId),Role="unknown",Profile=Config.Profile};
                        Solid shape=element is SpatialElement?SpatialPlan(r,"net","Проверка_ошибок"):null;
                        if(shape==null)foreach(var solid in Solids(element))shape=Union(shape,Projection(SolidUtils.CreateTransformed(solid,source.Transform)));
                        if(shape==null)continue;
                        current.Details.Add(new Detail{Metric="Проверка_ошибок",Source=source.Name,SourceKey=source.Key,Building=r.Building,Level=r.Level?.Name??"Без уровня",Elevation=r.Z,Element=issue.Element,UniqueId=element.UniqueId,
                            Purpose="Ошибка",Method=current.Method,Raw=shape.Volume*.09290304,Value=0,Unit="м²",Excluded=true,HasError=true,Reason=issue.Code+": "+issue.Message,Shape=shape});
                    }
                    catch(System.OperationCanceledException){throw;}
                    catch { /* An invalid element can lack even a diagnostic contour; its original structured issue remains. */ }
                }
            }
            private int DeleteOwned(bool graphics,bool schedules,Run run)
            {
                var all=new FilteredElementCollector(doc).WherePasses(new ExtensibleStorageFilter(StorageId)).Where(e=>Owned(e)&&
                    (graphics&&OneOf(Kind(e),"graphic","plan","3d","graphic-level","graphic-type")||schedules&&OneOf(Kind(e),"schedule","schedule-key"))).ToList();
                var allowed=new HashSet<long>(all.Select(e=>IDHelper.ElIdValue(e.Id)));int count=0;
                // First delete owned annotations and solids. Only then consider views/levels.
                foreach(var e in all.OrderBy(x=>x is View?1:x is Level?2:x is ElementType?3:0))
                {
                    if(doc.GetElement(e.Id)==null)continue;
                    using(var sub=new SubTransaction(doc))
                    {
                        sub.Start();var removed=doc.Delete(e.Id);
                        if(removed.Any(id=>!allowed.Contains(IDHelper.ElIdValue(id))))
                        {sub.RollBack();run.Issue("CLEANUP_DEPENDENCY","Предупреждение","Служебный объект сохранён: удаление затрагивает не принадлежащие расчёту элементы.",element:IDHelper.ElIdValue(e.Id).ToString(),action:"Проверьте добавленные пользователем аннотации / зависимости.");}
                        else {sub.Commit();count+=removed.Count;}
                    }
                }
                return count;
            }
            public int Cleanup()
            {var report=new Run();int count=0;Transaction("ТЭП: удалить служебные результаты",report,()=>count=DeleteOwned(true,true,report));if(Last!=null)Last.Issues.AddRange(report.Issues);return count;}
            private static Autodesk.Revit.DB.Color ParseColor(string text)
            {
                var c=(System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(text);return new Autodesk.Revit.DB.Color(c.R,c.G,c.B);
            }
            private OverrideGraphicSettings Override(Detail d,ElementId fill)
            {
                var colour=ParseColor(d.HasError?Config.ErrorColor:d.Excluded?Config.ExcludedColor:d.Manual?Config.ManualColor:
                    OneOf(d.Purpose,"balcony","loggia","terrace","veranda","cold-storage")?Config.SummerColor:
                    OneOf(d.Purpose,"public","nonresidential-common")?Config.PublicColor:
                    OneOf(d.Purpose,"technical-room","technical-space","technical-void","roof-vent")?Config.TechnicalColor:Config.IncludedColor);
                return new OverrideGraphicSettings().SetProjectionLineColor(colour).SetSurfaceForegroundPatternId(fill).SetSurfaceForegroundPatternColor(colour).SetSurfaceTransparency(d.VolumeShape?65:35);
            }
            public void CreateGraphics(Run run)
            {
                if(!ViewsEnabled)return;
                if(!run.Details.Any(d=>d.Shape!=null)){run.Issue("GRAPHICS_NO_GEOMETRY","Предупреждение","В сохранённом отчёте нет геометрии Revit. Выполните расчёт заново для построения видов.");return;}
                try
                {
                    Transaction("ТЭП: проверочные виды",run,()=>
                    {
                        if(Config.Graphics=="replace")DeleteOwned(true,false,run);
                        var fill=new FilteredElementCollector(doc).OfClass(typeof(FillPatternElement)).Cast<FillPatternElement>().FirstOrDefault(f=>f.GetFillPattern().IsSolidFill);
                        if(fill==null)throw new InvalidOperationException("В проекте отсутствует сплошная штриховка.");
                        var regionType=new FilteredElementCollector(doc).OfClass(typeof(FilledRegionType)).Cast<FilledRegionType>().FirstOrDefault();
                        if(regionType==null)throw new InvalidOperationException("В проекте отсутствует тип цветовой области (FilledRegionType).");
                        if(regionType.ForegroundPatternId!=fill.Id){regionType=(FilledRegionType)regionType.Duplicate("ТЭП_Заливка_"+run.Id.Substring(0,8));regionType.ForegroundPatternId=fill.Id;Tag(regionType,"graphic-type",run.Id);}
                        var planType=new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault(v=>v.ViewFamily==ViewFamily.FloorPlan);
                        var threeType=new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault(v=>v.ViewFamily==ViewFamily.ThreeDimensional);
                        var textType=new FilteredElementCollector(doc).OfClass(typeof(TextNoteType)).FirstOrDefault();
                        if(planType==null)throw new InvalidOperationException("В проекте нет типа плана этажа для проверочных обводок.");
                        var solidsByMetric=new Dictionary<string,List<ElementId>>();var views3d=new Dictionary<string,View3D>();
                        foreach(var metric in run.Details.Where(d=>d.Shape!=null).GroupBy(d=>d.Metric))
                        {
                            View3D v3=null;var ids=new List<ElementId>();solidsByMetric[metric.Key]=ids;
                            if(Config.Create3D&&threeType!=null){v3=View3D.CreateIsometric(doc,threeType.Id);v3.Name=UniqueName(Config.Prefix+metric.Key+"_3D");Tag(v3,"3d",run.Id);views3d[metric.Key]=v3;
                                var template=doc.GetElement(Config.View3DTemplate??"") as View;if(template!=null&&v3.IsValidViewTemplate(template.Id))v3.ViewTemplateId=template.Id;}
                            foreach(var floor in metric.GroupBy(d=>new {d.Building,d.Section,Z=Math.Round(d.Elevation,6)}))
                            {
                                ViewPlan plan=null;
                                if(planType!=null&&floor.Any(d=>!d.VolumeShape))
                                {
                                    var level=new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().FirstOrDefault(l=>Math.Abs(l.Elevation-floor.Key.Z)<1e-6);
                                    if(level==null){level=Level.Create(doc,floor.Key.Z);Tag(level,"graphic-level",run.Id);}
                                    plan=ViewPlan.Create(doc,planType.Id,level.Id);plan.Name=UniqueName(Config.Prefix+metric.Key+"_"+floor.Key.Building+"_"+floor.First().Level);Tag(plan,"plan",run.Id);plan.Scale=100;
                                    var template=doc.GetElement(Config.PlanTemplate??"") as View;if(template!=null&&plan.IsValidViewTemplate(template.Id))plan.ViewTemplateId=template.Id;
                                }
                                bool legend=false;
                                foreach(var detail in floor)
                                {
                                    Progress("Построение видов: "+detail.Metric+" / "+detail.Building+" / "+detail.Level+"; ID "+detail.Element);
                                    using(var sub=new SubTransaction(doc))
                                    {
                                        sub.Start();try
                                        {
                                            var overrides=Override(detail,fill.Id);var localShapes=new List<Solid>();
                                            if(detail.VolumeShape)localShapes.Add(detail.Shape);
                                            else
                                            {
                                                foreach(PlanarFace face in detail.Shape.Faces.Cast<Face>().OfType<PlanarFace>().Where(f=>f.FaceNormal.Z<-.999999))
                                                {
                                                    var loops=face.GetEdgesAsCurveLoops().Select(l=>FlatLoop(l,detail.Elevation)).ToList();
                                                    if(plan!=null){var region=FilledRegion.Create(doc,regionType.Id,plan.Id,loops);Tag(region,"graphic",run.Id+"|"+detail.SourceKey+"|"+detail.Element+"|"+detail.Metric);plan.SetElementOverrides(region.Id,overrides);detail.ViewId=plan.UniqueId;}
                                                    var measuredLoops=face.GetEdgesAsCurveLoops().Select(l=>FlatLoop(l,detail.MeasurementElevation??detail.Elevation)).ToList();
                                                    localShapes.Add(Extrude(measuredLoops,.1));
                                                }
                                            }
                                            if(v3!=null&&localShapes.Count>0)
                                            {var ds=DirectShape.CreateElement(doc,new ElementId(BuiltInCategory.OST_GenericModel));ds.ApplicationId=Owner;ds.ApplicationDataId=run.Id+"/"+detail.Metric+"/"+detail.SourceKey+"/"+detail.Element;
                                                ds.SetShape(localShapes.Cast<GeometryObject>().ToList());Tag(ds,"graphic",run.Id);ds.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)?.Set(detail.Metric+" | "+detail.Building+" | "+detail.Value.ToString("0.###")+" "+detail.Unit);v3.SetElementOverrides(ds.Id,overrides);ids.Add(ds.Id);if(detail.VolumeShape)detail.ViewId=v3.UniqueId;}
                                            if(plan!=null&&textType!=null)
                                            {
                                                var box=detail.Shape.GetBoundingBox();var p=box.Transform.OfPoint(box.Min);p=new XYZ(p.X,p.Y,detail.Elevation);
                                                string label=(detail.HasError?"ОШИБКА ":detail.Excluded?"ИСКЛЮЧЕНО ":detail.Manual?"КОРРЕКТИРОВКА ":"")+detail.Element+": "+detail.Raw.ToString("0.###")+" × "+detail.Factor.ToString("0.###")+" = "+detail.Value.ToString("0.###")+" "+detail.Unit;
                                                var note=TextNote.Create(doc,plan.Id,p,label,textType.Id);Tag(note,"graphic",run.Id);
                                                if(!legend){CreateLegend(plan,p+new XYZ(0,12,0),regionType.Id,textType.Id,fill.Id,run,detail);legend=true;}
                                            }
                                            sub.Commit();
                                        }
                                        catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){sub.RollBack();run.Issue("GRAPHIC_ELEMENT","Ошибка",ex.Message,detail.Metric,detail.Source,detail.Building,detail.Element,"Проверьте замкнутость контуров и ограничения вида Revit.");}
                                    }
                                }
                            }
                        }
                        // Hide only this tool's solids belonging to other indicators on its 3D views.
                        foreach(var pair in views3d)
                        {var other=solidsByMetric.Where(p=>p.Key!=pair.Key).SelectMany(p=>p.Value).Where(id=>doc.GetElement(id)!=null&&doc.GetElement(id).CanBeHidden(pair.Value)).ToList();if(other.Count>0)pair.Value.HideElements(other);}
                    });
                }
                catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){run.Issue("GRAPHICS_FAILED","Ошибка","Создание графики отменено целиком: "+ex.Message);foreach(var d in run.Details)d.ViewId=null;}
            }
            private void CreateLegend(ViewPlan view,XYZ origin,ElementId regionType,ElementId textType,ElementId fill,Run run,Detail detail)
            {
                var title=TextNote.Create(doc,view.Id,origin+new XYZ(0,3,0),"ТЭП | "+run.Method+" | "+run.Date+"\n"+detail.MetricName+" | "+detail.Building+" | "+detail.Level+"\nКонтур: до коэффициента; подпись и итог: после коэффициента.",textType);Tag(title,"graphic",run.Id);
                var entries=new[]{new Detail{Purpose="heated",Reason="Жилая / прочая включённая часть"},new Detail{Purpose="public",Reason="Общественная часть"},new Detail{Purpose="balcony",Reason="Летние помещения"},new Detail{Purpose="technical-room",Reason="Технические помещения"},new Detail{Excluded=true,Reason="Исключено / справочный контур"},new Detail{Manual=true,Reason="Ручное решение"},new Detail{HasError=true,Reason="Ошибка исходных данных"}};
                for(int i=0;i<entries.Length;i++)
                {
                    var p=origin-new XYZ(0,i*2.2,0);var points=new[]{p,p+new XYZ(2,0,0),p+new XYZ(2,1,0),p+new XYZ(0,1,0)};
                    var loop=CurveLoop.Create(Enumerable.Range(0,4).Select(j=>(Curve)Line.CreateBound(points[j],points[(j+1)%4])).ToList());
                    var swatch=FilledRegion.Create(doc,regionType,view.Id,new List<CurveLoop>{loop});Tag(swatch,"graphic",run.Id);view.SetElementOverrides(swatch.Id,Override(entries[i],fill));
                    var note=TextNote.Create(doc,view.Id,p+new XYZ(2.7,1,0),entries[i].Reason,textType);Tag(note,"graphic",run.Id);
                }
            }
            private void CreateSchedule(Run run)
            {
                if(!ViewsEnabled||!Config.CreateSchedule)return;
                try
                {
                    Transaction("ТЭП: сводная спецификация",run,()=>
                    {
                        if(Config.Graphics=="replace")DeleteOwned(false,true,run);
                        var schedule=ViewSchedule.CreateKeySchedule(doc,new ElementId(BuiltInCategory.OST_GenericModel));schedule.Name=UniqueName(Config.Prefix+"Сводные показатели");Tag(schedule,"schedule",run.Id);
                        schedule.KeyScheduleParameterName=UniqueName(Config.Prefix+"Ключ_"+run.Id.Substring(0,8));
                        var definition=schedule.Definition;
                        var fields=definition.GetSchedulableFields();
                        var comments=fields.FirstOrDefault(f=>IDHelper.ElIdValue(f.ParameterId)==(long)BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                        if(comments==null)throw new InvalidOperationException("Категория не поддерживает поле комментария для сводной спецификации.");
                        var field=definition.AddField(comments);field.ColumnHeading="Значение | единица | методика | статус";field.GridColumnWidth=2.2;
                        definition.GetField(0).ColumnHeading="Показатель";definition.GetField(0).GridColumnWidth=2.6;
                        doc.Regenerate();var body=schedule.GetTableData().GetSectionData(SectionType.Body);
                        foreach(var s in run.Summary)
                        {
                            var before=new HashSet<long>(new FilteredElementCollector(doc,schedule.Id).WhereElementIsNotElementType().Select(e=>IDHelper.ElIdValue(e.Id)));
                            int row=body.LastRowNumber+1;if(!body.CanInsertRow(row))row=body.FirstRowNumber;
                            body.InsertRow(row);doc.Regenerate();
                            var key=new FilteredElementCollector(doc,schedule.Id).WhereElementIsNotElementType().FirstOrDefault(e=>!before.Contains(IDHelper.ElIdValue(e.Id)));
                            if(key==null)throw new InvalidOperationException("Revit не вернул созданную строку ключевой спецификации.");
                            var name=key.get_Parameter(BuiltInParameter.REF_TABLE_ELEM_NAME);var value=key.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                            if(name==null||name.IsReadOnly||value==null||value.IsReadOnly)throw new InvalidOperationException("Параметры строки ключевой спецификации недоступны для записи.");
                            name.Set(s.Name);value.Set(s.ValueText(Config.Decimals)+" | "+s.Method+" | "+s.Status);Tag(key,"schedule-key",run.Id);
                        }
                    });
                }
                catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){run.Issue("SCHEDULE_FAILED","Ошибка","Сводная спецификация не создана; транзакция отменена: "+ex.Message);}
            }
            public void NavigateRequested()
            {
                var d=RequestedDetail;if(d==null)return;
                try
                {
                    if(!string.IsNullOrWhiteSpace(d.ViewId)){var view=doc.GetElement(d.ViewId) as View;if(view!=null)ui.ActiveView=view;}
                    var source=Sources.FirstOrDefault(s=>s.Key==d.SourceKey);ElementId id=source?.RootLink;
                    if(id==null&&source?.Document==doc)
                    {var element=string.IsNullOrWhiteSpace(d.UniqueId)?null:doc.GetElement(d.UniqueId);long numeric;if(element==null&&long.TryParse(d.Element,out numeric))element=doc.GetElement(IDHelper.CreateElementId(numeric));id=element?.Id;}
                    if(id!=null&&doc.GetElement(id)!=null){ui.Selection.SetElementIds(new List<ElementId>{id});ui.ShowElements(id);}
                }
                catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){TaskDialog.Show("ТЭП: переход к объекту",ex.Message);}
            }
        }
    }
}
