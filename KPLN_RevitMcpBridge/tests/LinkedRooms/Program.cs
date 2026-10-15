using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using KPLN_RevitMcpBridge.Server;
using KPLN_RevitMcpBridge.Services;

internal static class Program
{
    private sealed class Fixture
    {
        internal Document Host = new Document(), Linked = new Document();
        internal ViewPlan View;
        internal RevitLinkInstance Link;
        internal RevitLinkType LinkType;
        internal Level Level;
        internal Room Room;
        internal Dictionary<string, object> Row;
        internal int Count;
        internal Fixture()
        {
            Level = Linked.Add(new Level(), 10);
            Linked.Add(new Phase(), 20);
            var hostLevel = Host.Add(new Level(), 1);
            View = Host.Add(new ViewPlan { GenLevel=hostLevel, ViewType=ViewType.FloorPlan }, 2);
            View.Parameters[BuiltInParameter.VIEW_PHASE] = new Parameter { Value=new ElementId(3) };
            LinkType = Host.Add(new RevitLinkType(), 4);
            LinkType.Map[new ElementId(3)] = new ElementId(20);
            Link = Host.Add(new RevitLinkInstance { LinkedDocument=Linked, TypeId=LinkType.Id }, 5);
            var type = Host.Add(new RoomTagType(), 6);
            Row = new Dictionary<string, object> { {"view_unique_id",View.UniqueId}, {"link_instance_unique_id",Link.UniqueId}, {"tag_type_unique_id",type.UniqueId} };
            Room = AddRoom(30, 2, 3);
        }
        internal Room AddRoom(int id, double x, double y)
        {
            var room = Linked.Add(new Room { LevelId=Level.Id, Location=new LocationPoint { Point=new XYZ(x,y,0) },
                Category=new Category { Id=new ElementId(BuiltInCategory.OST_Rooms) } }, id);
            room.Parameters[BuiltInParameter.ROOM_PHASE_ID] = new Parameter { Value=new ElementId(20) };
            return room;
        }
        internal Dictionary<string, object> Run() => Json.Parse(Json.Serialize(LinkedRoomTagService.Create(Host, Row, ref Count)));
        internal RoomTag[] Tags => Host.Elements.OfType<RoomTag>().ToArray();
    }
    private static void Check(bool ok, string message) { if (!ok) throw new Exception(message); }
    private static void Reject(Action<Fixture> arrange, string code)
    {
        var f = new Fixture(); arrange(f);
        try { f.Run(); throw new Exception("Expected " + code); }
        catch (BridgeException e) { Check(e.Code == code, e.Code + " != " + code); }
        Check(f.Tags.Length == 0, "Validation must precede creation");
    }
    private static void Main()
    {
        var f = new Fixture();
        f.Link.Transform = new Transform { A=0, B=-1, C=1, D=0, X=100, Y=200, Z=10 };
        f.View.GenLevel.ProjectElevation = 10;
        f.Row["orientation"] = "Vertical";
        f.Run();
        Check(f.Tags.Single().TagHeadPosition.IsAlmostEqualTo(new XYZ(97,202,0)), "Rotated/translated room point");
        Check(f.Tags[0].TaggedRoomId.LinkInstanceId == f.Link.Id && f.Tags[0].TaggedRoomId.LinkedElementId == f.Room.Id, "Linked identity");
        Check(f.Tags[0].RoomTagType.UniqueId == (string)f.Row["tag_type_unique_id"] && !f.Tags[0].HasLeader, "Type and leader");
        Check(f.Tags[0].TagOrientation == SpatialElementTagOrientation.Vertical, "Orientation");
        f.Run(); Check(f.Tags.Length == 1, "Repeated batch must not duplicate tags");
        f.Tags[0].OwnerViewId = new ElementId(999); f.Run(); Check(f.Tags.Length == 2, "Other view is independent");
        f.Tags[1].TaggedRoomId.LinkInstanceId = new ElementId(999); f.Run(); Check(f.Tags.Length == 3, "Other link instance is independent");

        f = new Fixture(); f.Link.Transform.A=-1; f.Run();
        Check(f.Tags[0].TagHeadPosition.X == -2, "Mirrored link");
        f = new Fixture(); f.TagsHiddenTest();

        f = new Fixture();
        f.AddRoom(31, 3, 3).LevelId = new ElementId(11);
        f.AddRoom(32, 3, 3).Parameters[BuiltInParameter.ROOM_PHASE_ID].Value = new ElementId(21);
        f.AddRoom(33, 3, 3).Location = null;
        f.AddRoom(34, 3, 3).Area = 0;
        f.AddRoom(35, 200, 200);
        f.View.CropBoxActive = true;
        f.View.CropBox = new BoundingBoxXYZ { Min=new XYZ(0,0,100), Max=new XYZ(5,5,200) };
        var result = f.Run();
        Check(f.Tags.Length == 1, "Filter floor, phase, placement, area and crop; ignore crop Z");
        Check(Convert.ToInt32(result["matched_room_count"]) == 4, "Matched level/phase count");
        var skipped = (Dictionary<string, object>)result["skipped"];
        Check(Convert.ToInt32(skipped["unplaced_or_unenclosed"]) == 2 && Convert.ToInt32(skipped["outside_crop"]) == 1, "Skip report");

        f = new Fixture();
        f.View.CropBoxActive = true;
        f.View.CropBox = new BoundingBoxXYZ { Transform=new Transform { A=0,B=-1,C=1,D=0,X=10,Y=0 }, Min=new XYZ(0,0,0), Max=new XYZ(4,9,0) };
        f.Run(); Check(f.Tags.Length == 1, "Rotated crop in its own coordinates");

        Reject(x => x.Linked.Add(new Level(), 11), "ambiguous_linked_level");
        Reject(x => x.View.GenLevel.ProjectElevation=100, "ambiguous_linked_level");
        f = new Fixture(); f.View.GenLevel.ProjectElevation=100; f.Row["linked_level_unique_id"]=f.Level.UniqueId;
        f.Run(); Check(f.Tags.Length == 1, "Explicit level override");
        f = new Fixture(); f.View.ViewType=ViewType.CeilingPlan; f.Run(); Check(f.Tags.Length == 1, "Ceiling plan");
        Reject(x => x.Link.LinkedDocument=null, "link_unloaded");
        Reject(x => x.Row["tag_type_unique_id"]=x.View.UniqueId, "wrong_element_type");
        Reject(x => x.View.IsTemplate=true, "unsupported_view");
        Reject(x => x.View.ViewType=ViewType.AreaPlan, "unsupported_view");
        Reject(x => x.Link.Hidden=true, "hidden_annotation_source");
        Reject(x => x.View.TagsHidden=true, "hidden_annotation_source");
        Reject(x => x.LinkType.Map.Clear(), "missing_phase_mapping");
        Reject(x => x.Row["orientation"]="Diagonal", "invalid_input");
        Reject(x => x.Link.Transform.BasisZ=new XYZ(0,1,0), "unsupported_link_transform");
        Reject(x => { x.View.CropBoxActive=true; x.View.Crop.ShapeSet=true; }, "unsupported_crop");
        Reject(x => { x.View.CropBoxActive=true; x.View.Crop.NumberOfSplitRegions=2; }, "unsupported_crop");
        Reject(x => x.Count=1000, "too_many_tags");
        Reject(x => x.Host.ReturnNull=true, "tag_creation_failed");
        f = new Fixture(); f.Host.Orphan=true;
        try { f.Run(); throw new Exception("Orphan accepted"); }
        catch (BridgeException e) { Check(e.Code == "tag_creation_failed", "Orphan must fail transaction"); }
        Console.WriteLine("LinkedRooms: transform, scope, phase, crop, duplicates, limits and failures passed.");
    }
    private static void TagsHiddenTest(this Fixture f)
    {
        f.Run(); f.Tags[0].Hidden=true; f.Run(); Check(f.Tags.Length == 1, "Hidden existing tag must not duplicate");
        f.Tags[0].IsOrphaned=true; f.Run(); Check(f.Tags.Length == 2, "Orphan must not block replacement");
    }
}
