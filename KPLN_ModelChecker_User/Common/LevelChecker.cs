using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static KPLN_ModelChecker_User.Common.Collections;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_ModelChecker_User.Common
{
    public class LevelChecker
    {
        public static List<LevelChecker> LevelCheckers = new List<LevelChecker>();
        
        public double Min { get; }
        
        public double Max { get; }
        
        public Level Level { get; }
        
        public Level UpperLevel { get; private set; }
        
        private Document Doc { get; }

        public LevelChecker(Document doc, Level level, string code)
        {
            Doc = doc;
            Level = level;
            Min = level.Elevation;
            
            ElementId upperLevelId = level.get_Parameter(BuiltInParameter.LEVEL_UP_TO_LEVEL).AsElementId();
            
            if (upperLevelId == null || upperLevelId.IntegerValue == -1)
            {
                if (code != null)
                {
                    string name = Level.get_Parameter(BuiltInParameter.DATUM_TEXT).AsString();
                    string part = null;
                    if (name.Contains('_'))
                    {
                        part = name.Split('_')[0];
                    }
                    UpperLevel = GetNearestUpperLevel(level.Elevation, part);
                }
                else
                {
                    UpperLevel = GetNearestUpperLevel(level.Elevation, null);
                }
                if (UpperLevel != null)
                {
                    Max = UpperLevel.Elevation;
                }
                else
                {
                    Max = Min + 10000 / 304.8;
                }
            }
            else
            {
                UpperLevel = doc.GetElement(upperLevelId) as Level;
                Max = UpperLevel.Elevation;
            }
            //Print(string.Format("{0} ====> {1} ({2}mm:{3}mm)", Level.FieldName, UpperLevel != null ? UpperLevel.FieldName : "null", Math.Round(Min * 304.8, 2), Math.Round(Max * 304.8, 2)), KPLN_Loader.Preferences.MessageType.System_Regular);
        }

        public static CheckResult CheckLevels(Document doc)
        {
            HashSet<int> x = new HashSet<int>();
            foreach (Level level in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements())
            {
                string name = level.get_Parameter(BuiltInParameter.DATUM_TEXT).AsString();
                if (name.Contains('_'))
                {
                    string prefix = name.Split('_')[0];
                    if (prefix.StartsWith("С") && !(prefix.StartsWith("СТЛ")))
                    {
                        x.Add(0);
                    }
                    else
                    {
                        if (prefix.StartsWith("К") | prefix.StartsWith("ПАР") | prefix.StartsWith("СТЛ"))
                        {
                            x.Add(1);
                        }
                        else
                        {
                            if (prefix.StartsWith("C") || prefix.StartsWith("c") || prefix.StartsWith("K") || prefix.StartsWith("k"))
                            { Print(string.Format("Не допускается использование латинницы в наименовании уровней\nсм.уровень «{0}»", name), MessageType.Error); }
                            if (prefix.StartsWith("с") || prefix.StartsWith("к"))
                            { Print(string.Format("Некорректный регистр\nсм.уровень «{0}»", name), MessageType.Error); }
                            x.Add(2);
                        }
                    }
                }
                else
                {
                    return CheckResult.Error;
                }
            }
            if (x.Count != 1)
            {
                return CheckResult.Error;
            }
            else
            {
                if (x.Contains(0)) { return CheckResult.Sections; }
                if (x.Contains(1)) { return CheckResult.Corpus; }
                if (x.Contains(2)) { return CheckResult.NoSections; }
            }
            return CheckResult.Error;
        }
        
        public static void AddLevel(Level level, Document doc, string code)
        {
            LevelCheckers.Add(new LevelChecker(doc, level, code));
        }
        
        public static LevelChecker GetLevelById(ElementId id)
        {
            foreach (LevelChecker lvlChk in LevelCheckers)
            {
                if (lvlChk.Level.Id.IntegerValue == id.IntegerValue)
                {
                    return lvlChk;
                }
            }
            return null;
        }
        
        public static List<LevelChecker> GetOtherLevelById(ElementId id)
        {
            List<LevelChecker> levels = new List<LevelChecker>();
            foreach (LevelChecker lvlChk in LevelCheckers)
            {
                if (lvlChk.Level.Id.IntegerValue != id.IntegerValue)
                {
                    levels.Add(lvlChk);
                }
            }
            return levels;
        }
        
        public LevelCheckResult GetLevelIntersection(BoundingBoxXYZ box)
        {
            if (box.Min.Z > Max || box.Max.Z < Min) 
            {
                return LevelCheckResult.NotInside; 
            }
            if (box.Min.Z >= Min && box.Max.Z <= Max)
            {
                return LevelCheckResult.FullyInside;
            }
            else
            {
                double elementHeight = box.Max.Z - box.Min.Z;
                double max = Math.Min(box.Max.Z, Max);
                double min = Math.Max(box.Min.Z, Min);
                double intersectHeight = max - min;
                if (intersectHeight < elementHeight)
                { 
                    return LevelCheckResult.TheLeastInside;
                }
                else
                { 
                    return LevelCheckResult.MostlyInside;
                }
            }
        }

        public LevelCheckResult GetFloorLevelIntersection(BoundingBoxXYZ box)
        {
            if (box.Max.Z < Min - 1500 / 304.8 || box.Min.Z > Min + 1500 / 304.8)
            {
                return LevelCheckResult.NotInside;
            }
            if (box.Min.Z >= Min - 1500 / 304.8 && box.Max.Z <= Min + 1500 / 304.8)
            {
                return LevelCheckResult.FullyInside;
            }
            return LevelCheckResult.MostlyInside;
        }
        
        private List<Level> GetLevelsByP(double elevation, Document doc, double min, double max, string part = null)
        {
            List<Level> levels = new List<Level>();
            foreach (Element element in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements())
            {
                Level level = element as Level;
                if (part != null)
                {
                    if (!level.get_Parameter(BuiltInParameter.DATUM_TEXT).AsString().StartsWith(part))
                    {
                        continue;
                    }
                }
                if (level.Elevation > elevation && level.Elevation - elevation <= max / 304.8 && level.Elevation - elevation >= min / 304.8)
                {
                    levels.Add(level);
                }
            }
            return levels;
        }
        public Level GetNearestUpperLevel(double elevation, string part = null)
        {
            List<Level> levels = GetLevelsByP(elevation, Doc, 2000, 6000, part);
            if (levels.Count == 0)
            {
                levels = GetLevelsByP(elevation, Doc, 1000, 2000, part);
            }
            if (levels.Count == 0)
            {
                return null;
            }
            if(levels.Count == 1)
            {
                return levels[0];    
            }
            List<Level> sortedLevels = levels.OrderBy(o => o.Elevation).ToList();
            return sortedLevels[0];
        }
    }
}
