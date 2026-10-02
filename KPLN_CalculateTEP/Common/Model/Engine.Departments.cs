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
            public string Role { get; set; } = "unknown";
        }
        public class DepartmentRoom
        {
            public string Department { get; set; }
            public string Source { get; set; }
            public string Level { get; set; }
            public string Number { get; set; }
            public string Name { get; set; }
            public string Element { get; set; }
            public double Area { get; set; }
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
                if (Eq(value, "Квартира")) return "heated";
                // Manual choices for other departments take precedence on every rescan.
                var assigned = Config.Departments?.FirstOrDefault(d => Eq(d.Value, value));
                if (assigned != null) return assigned.Role;
                string role;
                return departmentRoleLabels.TryGetValue((value ?? "").Trim(), out role) ? role : "unknown";
            }
            public static string RoomPart(string department) { return Eq(department, "Квартира") ? "residential" : "nonresidential"; }
            public void AssignDepartment(string value, string role)
            {
                if (Eq(value, "Квартира") && role != "heated") throw new InvalidOperationException("Назначение «Квартира» закреплено за жилыми помещениями квартиры.");
                if (!Roles().Any(r => r.Key == role)) throw new InvalidOperationException("Неизвестная категория помещения.");
                Config.Departments = Config.Departments ?? new List<DepartmentAssignment>();
                Config.Departments.RemoveAll(d => Eq(d.Value, value));
                Config.Departments.Add(new DepartmentAssignment { Value = (value ?? "").Trim(), Role = role });
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
                var rooms = new List<DepartmentRoom>();
                foreach (var source in Sources.Where(s => s.Loaded && s.LoadError == null))
                {
                    int index = 0;
                    foreach (var room in source.Elements.OfType<Room>())
                    {
                        progress?.Invoke("Сканирование назначений: " + source.Name + "; помещений " + (++index));
                        if (!IsPlacedRoom(room) || !PhaseAccepted(source, room)) continue;
                        rooms.Add(new DepartmentRoom {
                            Department = Value(room, "@Department"), Source = source.Name,
                            Level = room.Level?.Name ?? "Уровень не определён", Number = room.Number, Name = room.Name,
                            Element = IDHelper.ElIdValue(room.Id).ToString(), Area = room.Area * .09290304
                        });
                    }
                }
                return GroupDepartments(rooms);
            }
        }
    }
}
