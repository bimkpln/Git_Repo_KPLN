using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace KPLN_Finishing.CommandTools
{
    public static class GroupParser
    {
        private static void LinkRoom(Element element, Room room)
        {
            Tools.ApplyRoom(element, room);
        }
        private static string GetKey(Element element, XYZ location)
        {
            int amm = 4;
            List<string> parts = new List<string>();
            parts.Add(element.GetTypeId().ToString());
            if (element.Location.GetType() == typeof(LocationCurve))
            {
                Curve curve = (element.Location as LocationCurve).Curve;
                Line l1 = Line.CreateBound(curve.GetEndPoint(0), location);
                parts.Add(Math.Round(l1.Length, amm).ToString());
                Line l2 = Line.CreateBound(curve.GetEndPoint(0), location);
                parts.Add(Math.Round(l2.Length, amm).ToString());
            }
            if (element.Location.GetType() == typeof(LocationPoint))
            {
                XYZ point = (element.Location as LocationPoint).Point;
                Line l3 = Line.CreateBound(point, location);
                parts.Add(Math.Round(l3.Length, amm).ToString());
            }
            return string.Join("_", parts);
        }
        private static string GetStrongKey(Element element, XYZ location)
        {
            int amm = 4;
            List<string> parts = new List<string>();
            parts.Add(element.GetTypeId().ToString());
            if (element.Location.GetType() == typeof(LocationCurve))
            {
                Curve curve = (element.Location as LocationCurve).Curve;
                Line l1 = Line.CreateBound(curve.GetEndPoint(0), location);
                parts.Add(Math.Round(l1.Length, amm).ToString());
                parts.Add(Math.Round(l1.Direction.X, amm).ToString());
                parts.Add(Math.Round(l1.Direction.Y, amm).ToString());
                parts.Add(Math.Round(l1.Direction.Z, amm).ToString());
                Line l2 = Line.CreateBound(curve.GetEndPoint(0), location);
                parts.Add(Math.Round(l2.Length, amm).ToString());
                parts.Add(Math.Round(l2.Direction.X, amm).ToString());
                parts.Add(Math.Round(l2.Direction.Y, amm).ToString());
                parts.Add(Math.Round(l2.Direction.Z, amm).ToString());
            }
            if (element.Location.GetType() == typeof(LocationPoint))
            {
                XYZ point = (element.Location as LocationPoint).Point;
                Line l3 = Line.CreateBound(point, location);
                parts.Add(Math.Round(l3.Length, amm).ToString());
                parts.Add(Math.Round(l3.Direction.X, amm).ToString());
                parts.Add(Math.Round(l3.Direction.Y, amm).ToString());
                parts.Add(Math.Round(l3.Direction.Z, amm).ToString());
            }
            return string.Join("_", parts);
        }
        public static void ParseGroup(Document doc, Group group)
        {
            Transaction t = new Transaction(doc, "Присоеденить типовые этажи");
            t.Start();
            if (group.Location == null) { throw new Exception("Группа не должна быть вложенной в другую группу!"); }
            foreach (ElementId member_id in group.GetMemberIds())
            {
                try
                {
                    Element finish_element = doc.GetElement(member_id);
                    Element finish_element_type = doc.GetElement(finish_element.GetTypeId());
                    if (finish_element_type.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString().ToLower() != "отделка")
                    {
                        continue;
                    }
                    Room room = null;
                    if (finish_element.Category.Id.IntegerValue != -2000011 && finish_element.Category.Id.IntegerValue != -2000038 && finish_element.Category.Id.IntegerValue != -2000032)
                    {
                        continue;
                    }
                    try
                    {
                        room = doc.GetElement(new ElementId(int.Parse(finish_element.LookupParameter(Names.parameter_Room_Id).AsString()))) as Room;
                        if (room == null)
                        {
                            continue;
                        }
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                    try
                    {
                        List<Wall> similar_wall_elements = new List<Wall>();
                        string wall_key = GetKey(finish_element, (group.Location as LocationPoint).Point);
                        string wall_room_key = GetKey(finish_element, (room.Location as LocationPoint).Point);
                        string wall_room_key_strong = GetStrongKey(finish_element, (room.Location as LocationPoint).Point);
                        List<Group> similar_groups = new List<Group>();
                        foreach (Group g in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_IOSModelGroups).WhereElementIsNotElementType().ToElements())
                        {
                            if (g.Location != null && g.Id.IntegerValue != group.Id.IntegerValue && g.GroupType.Id.IntegerValue == group.GroupType.Id.IntegerValue)
                            {
                                similar_groups.Add(g);
                            }
                        }
                        foreach (Group g in similar_groups)
                        {
                            XYZ g_location = (g.Location as LocationPoint).Point;
                            foreach (ElementId id in g.GetMemberIds())
                            {
                                Element element = doc.GetElement(id);
                                if (element.Category.Id.IntegerValue != -2000011 && element.Category.Id.IntegerValue != -2000038 && element.Category.Id.IntegerValue != -2000032)
                                {
                                    continue;
                                }
                                Element element_type = doc.GetElement(finish_element.GetTypeId());
                                if (element_type.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString().ToLower() != "отделка")
                                {
                                    continue;
                                }
                                if (element.GetType() == typeof(Wall))
                                {
                                    List<Room> _rooms = new List<Room>();
                                    List<Room> _rooms_strong = new List<Room>();
                                    Wall w = element as Wall;
                                    string w_key = GetKey(w, g_location);
                                    if (wall_key == w_key)
                                    {
                                        foreach (Room r in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements())
                                        {
                                            string w_room_key = GetKey(w, (r.Location as LocationPoint).Point);
                                            string w_room_key_strong = GetStrongKey(w, (r.Location as LocationPoint).Point);
                                            if (wall_room_key == w_room_key)
                                            {
                                                _rooms.Add(r);
                                            }
                                        }
                                    }
                                    if (_rooms.Count > 1)
                                    {
                                        if (_rooms_strong.Count != 0)
                                        {
                                            LinkRoom(w, _rooms_strong[0]);
                                        }
                                        else
                                        {
                                            if (_rooms.Count != 0)
                                            {
                                                LinkRoom(w, _rooms[0]);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (_rooms.Count != 0)
                                        {
                                            LinkRoom(w, _rooms[0]);
                                        }
                                        else
                                        {
                                            if (_rooms_strong.Count != 0)
                                            {
                                                LinkRoom(w, _rooms_strong[0]);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception)
                    { }
                }
                catch (Exception)
                { }
            }
            t.Commit();
        }
    }
}
