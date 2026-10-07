using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public class DepartmentAssignment
        {
            public string Value { get; set; }
            public string Scope { get; set; }
            public string Role { get; set; } = "unknown";
        }
        public class DepartmentRoom
        {
            public string Department { get; set; }
            public string RawDepartment { get; set; }
            public string RawName { get; set; }
            public string Source { get; set; }
            public string SourceKey { get; set; }
            public string UniqueId { get; set; }
            public string Level { get; set; }
            public string Number { get; set; }
            public string Name { get; set; }
            public string Element { get; set; }
            public double Area { get; set; }
            public bool IsApartment { get; set; }
        }
        public class DepartmentGroup
        {
            public string Value { get; set; }
            public string Label { get { return string.IsNullOrWhiteSpace(Value) ? "(Назначение не заполнено)" : Value; } }
            public List<DepartmentRoom> Rooms { get; set; }
            public int Count { get { return Rooms.Count; } }
            public string Caption { get { return Label + " (" + Count + ")"; } }
        }
        public partial class Engine
        {
            private static readonly Dictionary<string, string> departmentRoleLabels = Roles()
                .Where(r => r.Key != "unknown").ToDictionary(r => r.Label, r => r.Key, StringComparer.OrdinalIgnoreCase);

            public string DepartmentRole(string value)
            {
                // Manual choices for other departments take precedence on every rescan.
                var assigned = Config.Departments?.FirstOrDefault(d => Eq(d.Value, value) && (string.IsNullOrEmpty(d.Scope) || d.Scope == ClassificationScope));
                if (assigned != null) return assigned.Role;
                var tokens=(value??"").Split(new[]{" | "},2,StringSplitOptions.None);
                if(tokens.Length==2&&tokens[0]!="Семейство")
                {
                    string match=DictionaryRoomRole(tokens[0],tokens[1],Eq(tokens[0],"Квартира"));
                    if(match!="unknown")return match;
                }
                string lookup = value ?? "";
                foreach(var prefix in new[]{"Семейство | ","Квартира | "})if(lookup.StartsWith(prefix,StringComparison.Ordinal))lookup=lookup.Substring(prefix.Length);
                bool apartment = (value ?? "").StartsWith("Квартира | ",StringComparison.Ordinal);
                var entry = classificationDictionary.Entries.FirstOrDefault(e => Eq(e.Value,lookup) && (string.IsNullOrEmpty(e.Scope) || e.Scope=="apartment" && apartment));
                if(entry!=null)return entry.Role;
                if(Eq(value,"Квартира"))return "heated";
                string role;
                return departmentRoleLabels.TryGetValue(lookup.Trim(), out role) ? role : "unknown";
            }
            public static string RoomPart(string department) { return Eq(department, "Квартира") ? "residential" : "nonresidential"; }
            public void AssignDepartment(string value, string role)
            {
                if (!Roles().Any(r => r.Key == role)) throw new InvalidOperationException("Неизвестная категория помещения.");
                Config.Departments = Config.Departments ?? new List<DepartmentAssignment>();
                Config.Departments.RemoveAll(d => Eq(d.Value, value) && (string.IsNullOrEmpty(d.Scope) || d.Scope == ClassificationScope));
                Config.Departments.Add(new DepartmentAssignment { Value = (value ?? "").Trim(), Role = role, Scope = ClassificationScope });
            }
            public static List<DepartmentGroup> GroupDepartments(IEnumerable<DepartmentRoom> rooms)
            {
                return rooms.GroupBy(r => (r.Department ?? "").Trim(), StringComparer.OrdinalIgnoreCase)
                    .Select(g => new DepartmentGroup { Value = g.Key, Rooms = g.ToList() })
                    .OrderBy(g => g.Value, StringComparer.CurrentCultureIgnoreCase).ToList();
            }
            public List<DepartmentGroup> ScanDepartments(Action<string> progress = null)
            {
                // Read system ROOM_DEPARTMENT directly. No solids, contours, transactions or parameter guesses.
                phaseCache.Clear();
                LoadClassificationDictionary();
                var rooms = new List<DepartmentRoom>();
                foreach (var source in Sources.Where(s => s.Loaded && s.LoadError == null))
                {
                    int index = 0;
                    foreach (var room in source.Elements.Where(e=>e is Room))
                    {
                        progress?.Invoke("Сканирование назначений: " + source.Name + "; помещений " + (++index));
                        if (room is Room && !IsPlacedRoom((Room)room) || !PhaseAccepted(source, room)) continue;
                        string department;
                        try { department=ClassificationValue(room); } catch(Exception ex) { department="(Ошибка параметра: "+ex.Message+")"; }
                        rooms.Add(new DepartmentRoom {
                            Department = department, RawDepartment=Value(room,"@Department"),RawName=Value(room,"@Name"),Source = source.Name, SourceKey=source.Key, UniqueId=room.UniqueId,IsApartment=IsApartmentRoom(room),
                            Level = (room.Document.GetElement(room.LevelId) as Level)?.Name ?? "Уровень не определён", Number = (room as Room)?.Number, Name = room.Name,
                            Element = IDHelper.ElIdValue(room.Id).ToString(), Area = (room as Room)?.Area * .09290304 ?? 0
                        });
                    }
                }
                ScanFamilyChoices(progress);
                return GroupDepartments(rooms);
            }
            private bool IsApartmentRoom(Element room)
            {
                if(Eq(Value(room,"@Department"),"Квартира"))return true;
                try{string number=Mapped(room,"apartment");return !string.IsNullOrWhiteSpace(number)&&number!="0"&&number!="-";}
                catch{return false;}
            }
        }
    }
}
