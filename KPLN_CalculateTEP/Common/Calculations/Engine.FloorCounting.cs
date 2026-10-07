using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public partial class Engine
        {
            // Read an explicitly written floor number, not a building suffix or any number
            // found in a datum name. No project-specific K1/K2 or source-name rules.
            public static string FloorCountDesignation(string name)
            {
                if(string.IsNullOrWhiteSpace(name))return null;
                const string number=@"(?<n>[+\-−]?\s*\d+)(?<letter>[а-яa-z]?)";
                const string marker=@"(?:этаж|эт\.?)(?![\p{L}])";
                var values=new List<string>();
                foreach(string pattern in new[]{@"(?<![\p{L}\p{N}.,])"+number+@"(?:-(?:й|ый|ой))?[\s_]*"+marker+@"(?!\p{N})",
                    @"(?<![\p{L}\p{N}])"+marker+@"[\s_]*"+number+@"(?![\p{L}\p{N}.,])"})
                    foreach(Match match in Regex.Matches(name,pattern,RegexOptions.IgnoreCase|RegexOptions.CultureInvariant))
                    {
                        int first=match.Groups["n"].Index,last=match.Groups["letter"].Index+match.Groups["letter"].Length;
                        string token=Regex.Replace(match.Groups["n"].Value.Replace('−','-'),@"\s+","");
                        if(Regex.IsMatch(name.Substring(0,first),@"\d\s*[-−–—/]\s*$")||
                            token.StartsWith("-",StringComparison.Ordinal)&&Regex.IsMatch(name.Substring(0,first),@"\d\s*$")||
                            Regex.IsMatch(name.Substring(last),@"^\s*[-−–—/]\s*[+\-−]?\d"))return null;
                        int ordinal;
                        if(int.TryParse(token,NumberStyles.AllowLeadingSign,CultureInfo.InvariantCulture,out ordinal))
                            values.Add(ordinal.ToString(CultureInfo.InvariantCulture)+match.Groups["letter"].Value.ToUpperInvariant());
                    }
                var distinct=values.Distinct(StringComparer.Ordinal).Take(2).ToList();
                return distinct.Count==1?distinct[0]:null;
            }

            public static string[] ResolveFloorCountKeys(string[] names,double[] elevations)
            {
                if(names==null||elevations==null||names.Length!=elevations.Length||elevations.Any(z=>double.IsNaN(z)||double.IsInfinity(z)))
                    throw new ArgumentException("Некорректный список уровней для подсчёта этажей.");
                const double tolerance=1/304.8; // 1 mm, only for coincident Revit datums.
                var designations=names.Select(FloorCountDesignation).ToArray();
                var keys=designations.Select(d=>d==null?null:"floor:"+d).ToArray();
                var unresolved=new List<int>();
                for(int i=0;i<names.Length;i++)
                {
                    if(designations[i]!=null)continue;
                    var matching=Enumerable.Range(0,names.Length).Where(j=>designations[j]!=null&&Math.Abs(elevations[i]-elevations[j])<=tolerance)
                        .Select(j=>keys[j]).Distinct(StringComparer.Ordinal).Take(2).ToList();
                    if(matching.Count==1)keys[i]=matching[0];
                    else if(matching.Count==0)unresolved.Add(i);
                    // More than one known floor at this elevation is ambiguous. A service
                    // datum must not become an additional floor or join those two floors.
                }
                double start=double.NaN;string key=null;
                foreach(int i in unresolved.OrderBy(i=>elevations[i]).ThenBy(i=>names[i],StringComparer.Ordinal))
                {
                    if(double.IsNaN(start)||elevations[i]-start>tolerance)
                    {start=elevations[i];key="elevation:"+start.ToString("R",CultureInfo.InvariantCulture);}
                    keys[i]=key;
                }
                return keys;
            }

            private List<List<Record>> GroupFloorCountRecords(List<Record> physical,string metric)
            {
                var result=new List<List<Record>>();
                foreach(var owner in physical.Where(r=>r.Level!=null).GroupBy(r=>FloorCountIdentity(r.Building,r.Section,0)))
                {
                    var levels=owner.GroupBy(r=>r.Level.Key,StringComparer.Ordinal).Select(g=>g.ToList()).ToList();
                    var keys=ResolveFloorCountKeys(levels.Select(g=>g[0].Level.Name).ToArray(),levels.Select(g=>g[0].Z).ToArray());
                    var groups=new Dictionary<string,List<Record>>(StringComparer.Ordinal);
                    for(int i=0;i<levels.Count;i++)
                    {
                        if(keys[i]==null)
                        {
                            Notice("FLOOR_LEVEL_AMBIGUOUS","Ошибка","Уровень «"+levels[i][0].Level.Name+
                                "» совпадает по отметке с несколькими этажами с разными номерами. Он не посчитан отдельным этажом; уточните обозначение уровня.",levels[i][0],metric);
                            continue;
                        }
                        List<Record> group;
                        if(!groups.TryGetValue(keys[i],out group))groups[keys[i]]=group=new List<Record>();
                        group.AddRange(levels[i]);
                    }
                    result.AddRange(groups.Values);
                }
                return result;
            }
        }
    }
}
