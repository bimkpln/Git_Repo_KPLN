using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public partial class Engine
        {
            public static string ResolveAssignedBuilding(IEnumerable<BuildingAssignment> rules,string source,string level,string group,Func<string,string> parameter)
            {
                var matches=new List<string>();
                foreach(var rule in rules??Enumerable.Empty<BuildingAssignment>())
                {
                    if(rule==null||string.IsNullOrWhiteSpace(rule.Match)||string.IsNullOrWhiteSpace(rule.Building))continue;
                    string value=rule.Scope=="source"?source:rule.Scope=="level"?level:rule.Scope=="group"?group:rule.Scope=="parameter"&&!string.IsNullOrWhiteSpace(rule.Parameter)?parameter(rule.Parameter):null;
                    if(value==null)continue;
                    bool match=rule.Contains?value.IndexOf(rule.Match.Trim(),StringComparison.OrdinalIgnoreCase)>=0:string.Equals(value.Trim(),rule.Match.Trim(),StringComparison.OrdinalIgnoreCase);
                    if(match)matches.Add(rule.Building.Trim());
                }
                var buildings=matches.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                if(buildings.Count>1)throw new InvalidOperationException("Конфликт соответствий корпусов: "+string.Join(", ",buildings)+". Уточните условия правил; корпус не назначен.");
                return buildings.FirstOrDefault();
            }
            private string AssignedBuilding(Source source,Element element)
            {
                return ResolveAssignedBuilding(Config.BuildingAssignments,source?.Name??element.Document.Title,ModelLevel(element)?.Name,
                    element is Group?element.Name:null,name=>Value(element,name));
            }
            public InputAudit InputAuditForRun {get;set;}
        }
    }
}
