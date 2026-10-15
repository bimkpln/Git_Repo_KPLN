using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using KPLN_RevitMcpBridge.Services;
using KPLN_RevitMcpBridge.Server;

class Program
{
    static int count;
    static XYZ P(double x,double y,double z=0) => new XYZ(x/304.8,y/304.8,z/304.8);
    static Outline O(double x,double y,double xx,double yy) => new Outline(P(x,y),P(xx,yy));
    static void Assert(bool value,string name) { if(!value)throw new Exception(name); count++; }
    static void Reject(Action action,string code) { try { action(); } catch(BridgeException ex) { Assert(ex.Code==code,ex.Code+" != "+code); return; } throw new Exception("Expected "+code); }
    static Dictionary<string,object> Input() => new Dictionary<string,object>{{"sheet_unique_id","sheet"}};
    static Dictionary<string,object> Move(string uid,double[] from,double[] to) => new Dictionary<string,object>{{"sheet_unique_id","sheet"},{"instance_unique_id",uid},{"expected_point_mm",from},{"point_mm",to}};
    static void Main()
    {
        CheckLabelMovement();
        // Смещение вдоль направления взгляда не меняет печатную позицию марки.
        var normal=new XYZ(0,0,-1);
        Assert(AnnotationPlane.Distance(P(1,2,4100),P(1,2,3955),normal)<1e-9,"Pipe leader depth normalization");
        Assert(AnnotationPlane.Distance(P(1,2,4100),P(1,2,0),normal)<1e-9,"Linked room native depth");
        Assert(Math.Abs(AnnotationPlane.Distance(P(1,2,0),P(1.2,2,100),normal)*304.8-.2)<1e-9,"Real planar shift is detected");
        Assert(AnnotationPlane.KeepDepth(P(1,2,4100),P(8,9,0),normal).DistanceTo(P(1,2,0))<1e-9,"Projection preserves both requested coordinates");
        var sectionNormal=new XYZ(1,0,0);
        Assert(AnnotationPlane.Distance(P(0,2,3),P(100,2,3),sectionNormal)<1e-9,"Section depth normalization");
        Assert(Math.Abs(AnnotationPlane.Distance(P(0,2,3),P(100,2,3.2),sectionNormal)*304.8-.2)<1e-9,"Section vertical shift is detected");
        var diagonal=new XYZ(Math.Sqrt(.5),Math.Sqrt(.5),0);
        Assert(AnnotationPlane.Distance(P(1,2,3),P(11,12,3),diagonal)<1e-9,"Rotated view normal");
        Assert(AnnotationPlane.Distance(P(1,2,3),P(11,-8,3),diagonal)>.01,"Rotated view planar shift");
        var doc=new Document();
        var sheet=new ViewSheet{Id=new ElementId(1),UniqueId="sheet",Outline=new BoundingBoxUV{Min=new UV(-841/304.8,0),Max=new UV(0,594/304.8)}};
        var type=new Element{Id=new ElementId(2),UniqueId="type",Label=new Parameter{Value=0}};
        var vp=new Viewport{Id=new ElementId(3),UniqueId="plan",OwnerViewId=sheet.Id,TypeId=type.Id,Bounds=O(-779.964,143.298,-140.036,546.702)};
        var schedule=new ScheduleSheetInstance{Id=new ElementId(4),UniqueId="schedule",OwnerViewId=sheet.Id};
        schedule.Point=P(-200,560);
        schedule.Box=new BoundingBoxXYZ{Min=P(-200,386),Max=P(-73,560)};
        sheet.Viewports.Add(vp.Id); doc.Elements.AddRange(new Element[]{sheet,type,vp,schedule});

        // Регрессия листа ТРЦ: оси вида пересекают таблицу; размеры в мм листа.
        var report=SheetLayoutService.Check(doc,Input());
        Assert(report.complete&&!report.bounds_clear&&report.issues.Length==1,"TRC schedule overlap must be detected");
        Reject(()=>SheetLayoutService.ValidateSheets(doc,new[]{"sheet"}),"sheet_layout_conflict");
        Assert(!report.frame_supplied&&report.visual_review_required,"No false frame/visual approval");

        // Таблица переносится в свободное поле с сохранением идентификатора.
        var moved=SheetLayoutService.Move(doc,Move("schedule",new[]{-200d,560d,0d},new[]{-200d,320d,0d}));
        Assert(ReferenceEquals(moved,schedule)&&schedule.UniqueId=="schedule","Identity preserved");
        Assert(SheetLayoutService.Check(doc,Input()).bounds_clear==false,"Moving down still overlaps plan");
        SheetLayoutService.Move(doc,Move("schedule",new[]{-200d,320d,0d},new[]{-139d,560d,0d}));
        Assert(SheetLayoutService.Check(doc,Input()).bounds_clear,"Separated layout");
        var input=Input(); input["minimum_gap_mm"]=2d;
        Assert(!SheetLayoutService.Check(doc,input).bounds_clear,"One mm clearance is insufficient");
        input["minimum_gap_mm"]=0d;
        input["content_min_mm"]=new[]{-821d,5d,0d}; input["content_max_mm"]=new[]{-20d,589d,0d};
        Assert(!SheetLayoutService.Check(doc,input).bounds_clear,"Content crossing frame must fail");
        input["content_max_mm"]=new[]{-5d,589d,0d};
        Assert(SheetLayoutService.Check(doc,input).bounds_clear,"Inside supplied frame");
        input["reserved_regions"]=new object[]{new Dictionary<string,object>{{"name","stamp"},{"min_mm",new[]{-190d,0d,0d}},{"max_mm",new[]{-5d,590d,0d}}}};
        Assert(!SheetLayoutService.Check(doc,input).bounds_clear,"Stamp reserved field protected");

        Reject(()=>SheetLayoutService.Move(doc,Move("schedule",new[]{-200d,560d,0d},new[]{-100d,400d,0d})),"stale_position");
        schedule.Pinned=true;
        Reject(()=>SheetLayoutService.Move(doc,Move("schedule",new[]{-139d,560d,0d},new[]{-100d,400d,0d})),"protected_element");
        schedule.Pinned=false; schedule.GroupId=new ElementId(5);
        Reject(()=>SheetLayoutService.Move(doc,Move("schedule",new[]{-139d,560d,0d},new[]{-100d,400d,0d})),"protected_element");
        schedule.GroupId=ElementId.InvalidElementId; schedule.IsTitleblockRevisionSchedule=true;
        Reject(()=>SheetLayoutService.Move(doc,Move("schedule",new[]{-139d,560d,0d},new[]{-100d,400d,0d})),"protected_element");
        schedule.Box=null;
        Assert(SheetLayoutService.Check(doc,Input()).bounds_clear,"Embedded revision schedule ignored");
        schedule.IsTitleblockRevisionSchedule=false;
        report=SheetLayoutService.Check(doc,Input());
        Assert(!report.complete&&!report.bounds_clear&&report.unreadable.Length==1,"Missing bbox never passes");
        schedule.OwnerViewId=new ElementId(999);
        Reject(()=>SheetLayoutService.Move(doc,Move("schedule",new[]{-139d,560d,0d},new[]{-100d,400d,0d})),"wrong_sheet");
        schedule.OwnerViewId=sheet.Id;
        schedule.Box=new BoundingBoxXYZ{Min=P(0,0),Max=P(10,20),Transform=new Transform{Origin=P(-100,570),Rotate=true}};
        var transformed=SheetLayoutService.BoundingRegion(schedule,sheet);
        Assert(Math.Abs(transformed.min_mm[0]+120)<.001&&Math.Abs(transformed.max_mm[1]-580)<.001,"Transformed eight-corner bounds");

        type.Label.Value=1; vp.LabelBounds=O(-150,567,-70,580);
        Assert(!SheetLayoutService.Check(doc,Input()).bounds_clear,"Viewport title overlap detected separately");
        vp.LabelBounds=null;
        Assert(!SheetLayoutService.Check(doc,Input()).complete,"Unreadable visible label is incomplete");
        type.Label.Value=0;
        var center=SheetLayoutService.PlacementPoint(vp);
        Assert(center.Z==0,"Ignore auxiliary viewport Z");
        SheetLayoutService.Move(doc,Move("plan",new[]{center.X*304.8,center.Y*304.8,0},new[]{-450d,340d,0d}));
        Assert(Math.Abs(SheetLayoutService.PlacementPoint(vp).X*304.8+450)<.001,"Viewport moves with same API");
        vp.IgnoreMove=true;
        Reject(()=>SheetLayoutService.Move(doc,Move("plan",new[]{-450d,340d,0d},new[]{-440d,340d,0d})),"placement_failed");
        Reject(()=>SheetLayoutService.Move(doc,Move("plan",new[]{-450d,340d,0d},new[]{-440d,340d,1d})),"invalid_point");
        Reject(()=>SheetLayoutService.Move(doc,Move("plan",new[]{-450d,340d,0d},new[]{double.NaN,340d,0d})),"invalid_value");
        input=Input();input["content_min_mm"]=new[]{-821d,5d,0d};
        Reject(()=>SheetLayoutService.Check(doc,input),"invalid_point");
        input=Input();input["minimum_gap_mm"]=-1d;
        Reject(()=>SheetLayoutService.Check(doc,input),"invalid_input");
        // Увеличение бумаги само по себе не перемещает элементы: после него проверяем новую композицию.
        doc.Elements.Remove(vp);sheet.Viewports.Clear();
        schedule.Box=new BoundingBoxXYZ{Min=P(10,300),Max=P(137,474)};
        Assert(!SheetLayoutService.Check(doc,Input()).bounds_clear,"Oversize composition rejected");
        sheet.Outline.Max=new UV(150/304.8,594/304.8);
        Assert(SheetLayoutService.Check(doc,Input()).bounds_clear,"Larger verified sheet can contain composition");
        schedule.Box=new BoundingBoxXYZ{Min=P(10,300),Max=P(10,300)};
        Assert(!SheetLayoutService.Check(doc,Input()).complete,"Empty schedule bounds are not safe");
        Console.WriteLine("PASS: "+count+" sheet layout and movement checks");
    }

    static void CheckLabelMovement()
    {
        var doc = new Document();
        var sheet = new ViewSheet { Id=new ElementId(100), UniqueId="label-sheet", Outline=new BoundingBoxUV { Min=new UV(0,0), Max=new UV(841/304.8,594/304.8) } };
        var type = new Element { Id=new ElementId(101), UniqueId="label-type", Label=new Parameter { Value=1 } };
        var vp = new Viewport { Id=new ElementId(102), UniqueId="label-plan", OwnerViewId=sheet.Id, TypeId=type.Id,
            Bounds=O(30,100,550,450), LabelBounds=O(30,80,200,95), LabelLineLength=170/304.8, Rotation=1 };
        doc.Elements.AddRange(new Element[] { sheet,type,vp });
        sheet.Viewports.Add(vp.Id);
        var row = new Dictionary<string,object> { {"sheet_unique_id",sheet.UniqueId}, {"viewport_unique_id",vp.UniqueId},
            {"expected_label_offset_mm",new[]{0d,0d,0d}}, {"label_offset_mm",new[]{200d,400d,0d}} };
        var center = vp.GetBoxCenter();
        var result = SheetLayoutService.UpdateLabel(doc,row);
        Assert(ReferenceEquals(result,vp) && vp.GetBoxCenter().DistanceTo(center)<1e-9 && vp.Rotation==1,"Label only: viewport identity, center and rotation preserved");
        Assert(vp.LabelOffset.DistanceTo(P(200,400))<1e-9 && Math.Abs(vp.LabelLineLength*304.8-170)<1e-9,"Offset in paper mm; omitted length preserved");
        SheetLayoutService.ValidateSheets(doc,new[]{sheet.UniqueId});
        Reject(()=>SheetLayoutService.UpdateLabel(doc,row),"stale_position");

        row["expected_label_offset_mm"]=new[]{200d,400d,0d};
        row["label_line_length_mm"]=-120d;
        Reject(()=>SheetLayoutService.UpdateLabel(doc,row),"invalid_input");
        row["expected_label_line_length_mm"]=180d;
        Reject(()=>SheetLayoutService.UpdateLabel(doc,row),"stale_position");
        row["expected_label_line_length_mm"]=170d;
        SheetLayoutService.UpdateLabel(doc,row);
        Assert(Math.Abs(vp.LabelLineLength*304.8+120)<1e-9,"Signed line length is supported");
        row.Remove("label_line_length_mm"); row.Remove("expected_label_line_length_mm");
        vp.Pinned=true;
        Reject(()=>SheetLayoutService.UpdateLabel(doc,row),"protected_element"); vp.Pinned=false;
        vp.GroupId=new ElementId(9);
        Reject(()=>SheetLayoutService.UpdateLabel(doc,row),"protected_element"); vp.GroupId=ElementId.InvalidElementId;
        vp.OwnerViewId=new ElementId(9);
        Reject(()=>SheetLayoutService.UpdateLabel(doc,row),"wrong_sheet"); vp.OwnerViewId=sheet.Id;
        sheet.IsPlaceholder=true;
        Reject(()=>SheetLayoutService.UpdateLabel(doc,row),"wrong_sheet"); sheet.IsPlaceholder=false;
        row["label_offset_mm"]=new[]{1d,2d,1d};
        Reject(()=>SheetLayoutService.UpdateLabel(doc,row),"invalid_point");
        row["label_offset_mm"]=new[]{double.NaN,0d,0d};
        Reject(()=>SheetLayoutService.UpdateLabel(doc,row),"invalid_value");
        row["label_offset_mm"]=new[]{250d,400d,0d}; vp.IgnoreLabelMove=true;
        Reject(()=>SheetLayoutService.UpdateLabel(doc,row),"placement_failed"); vp.IgnoreLabelMove=false;
        row["label_offset_mm"]=new[]{800d,400d,0d};
        SheetLayoutService.UpdateLabel(doc,row);
        Reject(()=>SheetLayoutService.ValidateSheets(doc,new[]{sheet.UniqueId}),"sheet_layout_conflict");
    }
}
