using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public partial class Engine
        {
            public static List<string> UpgradeSettings(Settings settings)
            {
                if(settings==null)throw new InvalidOperationException("В RVT нет содержимого настроек ТЭП.");
                if(settings.Version!=2&&settings.Version!=Settings.CurrentVersion)
                    throw new InvalidOperationException("Версия настроек "+settings.Version+" не поддерживается; этот плагин читает версии 2 - "+Settings.CurrentVersion+".");
                var notes=new List<string>();
                if(settings.Version<Settings.CurrentVersion)notes.Add("Настройки версии "+settings.Version+" приведены к версии "+Settings.CurrentVersion+".");
                var defaults=new Settings();
                // JSON from an older assembly may omit newly introduced fields. Only missing values receive defaults.
                foreach(var property in typeof(Settings).GetProperties().Where(p=>p.CanRead&&p.CanWrite&&!p.PropertyType.IsValueType))
                    if(property.GetValue(settings)==null&&property.GetValue(defaults)!=null)property.SetValue(settings,property.GetValue(defaults));
                int removed=0;
                foreach(var choice in new[]{Tuple.Create("Method","method"),Tuple.Create("Profile","profile"),Tuple.Create("Grouping","group"),Tuple.Create("Geometry","geometry"),Tuple.Create("Graphics","graphics"),Tuple.Create("ApartmentMode","count")})
                {
                    var property=typeof(Settings).GetProperty(choice.Item1);var value=(string)property.GetValue(settings);
                    if(!Choices(choice.Item2).Any(c=>c.Key==value)){property.SetValue(settings,property.GetValue(defaults));removed++;}
                }
                var metrics=Catalog().ToDictionary(m=>m.Key);
                var roles=new HashSet<string>(Roles().Select(r=>r.Key));
                var parameters=new HashSet<string>(Settings.DefaultParameters().Select(p=>p.Key));
                int before=settings.Metrics.Count;
                settings.Metrics=settings.Metrics.Where(m=>m!=null&&m.Key!=null&&metrics.ContainsKey(m.Key)).GroupBy(m=>m.Key).Select(g=>
                {
                    if(g.Count()==1)return g.First();
                    var metric=metrics[g.Key];metric.Enabled=g.All(m=>m.Enabled);return metric;
                }).ToList();removed+=before-settings.Metrics.Count;
                foreach(var metric in settings.Metrics)
                {
                    // These controls were removed from the UI; current calculation policy owns their values.
                    metric.ContourMode=RoomEnvelopeTotal((Indicator)Enum.Parse(typeof(Indicator),metric.Key))?"walls":"current";
                    metric.WallSelectionMode="exterior";metric.WallBoundary="auto";metric.WallCutHeight="0";
                    metric.ContourExclusions="normative";metric.VolumeMaskHeight="actual";metric.SelectedWalls=new List<string>();
                }
                before=settings.Parameters.Count;
                settings.Parameters=settings.Parameters.Where(p=>p!=null&&p.Key!=null&&parameters.Contains(p.Key)).GroupBy(p=>p.Key).Select(g=>
                {var p=g.First();if(g.Select(x=>x.Name??"").Distinct().Count()>1)p.Name="";return p;}).ToList();removed+=before-settings.Parameters.Count;
                foreach(var p in settings.Parameters.Where(p=>Settings.IsFixedParameter(p.Key)))p.Name=Settings.FixedParameterName(p.Key);
                before=settings.Departments.Count;
                settings.Departments=settings.Departments.Where(d=>d!=null&&d.Role!=null&&roles.Contains(d.Role)).GroupBy(d=>(d.Value??"").Trim(),StringComparer.OrdinalIgnoreCase)
                    .Where(g=>g.Select(d=>d.Role).Distinct().Count()==1).Select(g=>g.First()).ToList();removed+=before-settings.Departments.Count;
                before=settings.CategorySources.Count;
                settings.CategorySources=settings.CategorySources.Where(c=>c!=null&&c.Role!=null&&roles.Contains(c.Role)&&!(c.Role=="unknown"&&c.Families))
                    .GroupBy(c=>c.Role).Where(g=>g.Select(c=>c.Families).Distinct().Count()==1).Select(g=>g.First()).ToList();removed+=before-settings.CategorySources.Count;
                before=settings.CategoryFamilies.Count;
                settings.CategoryFamilies=settings.CategoryFamilies.Where(c=>c!=null&&c.Role!=null&&roles.Contains(c.Role)&&c.Role!="unknown"&&!string.IsNullOrWhiteSpace(c.Family)&&!string.IsNullOrWhiteSpace(c.Type))
                    .GroupBy(c=>c.Caption,StringComparer.OrdinalIgnoreCase).Where(g=>g.Select(c=>c.Role+"/"+c.AreaFromParameter+"/"+c.AreaParameter).Distinct().Count()==1).Select(g=>g.First()).ToList();removed+=before-settings.CategoryFamilies.Count;
                before=settings.StairFamilies.Count;
                settings.StairFamilies=settings.StairFamilies.Where(s=>s!=null&&!string.IsNullOrWhiteSpace(s.Family)&&!string.IsNullOrWhiteSpace(s.Type))
                    .GroupBy(s=>s.Caption,StringComparer.OrdinalIgnoreCase).Where(g=>g.Select(s=>s.WholeStair+"/"+s.WidthAxis+"/"+s.WidthParameter).Distinct().Count()==1).Select(g=>g.First()).ToList();removed+=before-settings.StairFamilies.Count;
                before=settings.Levels.Count;
                settings.Levels=new ObservableCollection<LevelSetting>(settings.Levels.Where(l=>l!=null&&!string.IsNullOrWhiteSpace(l.Key)&&string.IsNullOrWhiteSpace(l.Building)&&string.IsNullOrWhiteSpace(l.Section))
                    .GroupBy(l=>l.Key).Where(g=>g.Select(Serialize).Distinct().Count()==1).Select(g=>g.First()));removed+=before-settings.Levels.Count;
                foreach(var level in settings.Levels)
                {
                    if(!Choices("level").Any(c=>c.Key==level.Kind)){level.Kind="normal";removed++;}
                    if(!Choices("above").Any(c=>c.Key==level.Above)){level.Above="auto";removed++;}
                }
                before=settings.Sources.Count;
                settings.Sources=settings.Sources.Where(s=>s!=null&&!string.IsNullOrWhiteSpace(s.Key)&&Choices("source").Any(c=>c.Key==s.Mode))
                    .GroupBy(s=>s.Key).Where(g=>g.Select(s=>s.Mode+"/"+s.Profile).Distinct().Count()==1).Select(g=>g.First()).ToList();removed+=before-settings.Sources.Count;
                removed+=settings.Rules.Count+settings.Buildings.Count+settings.Corrections.Count+settings.Contours.Count;
                settings.Rules.Clear();settings.Buildings.Clear();settings.Corrections.Clear();settings.Contours.Clear();settings.PreviousWorkflowSettings=null;
                if(removed>0)notes.Add("Очищены устаревшие или конфликтующие настройки: "+removed+". Остальные значения сохранены.");
                settings.Version=Settings.CurrentVersion;
                return notes;
            }

            internal static int RemoveMissingSettingsReferences(Settings settings,ISet<string> sourceKeys,ISet<string> loadedSourceKeys,ISet<string> levelKeys)
            {
                Func<string,bool> missing=key=>
                {
                    if(string.IsNullOrWhiteSpace(key))return false;
                    int split=key.LastIndexOf('/');if(split<0)return true;
                    string source=key.Substring(0,split);
                    return !sourceKeys.Contains(source)||loadedSourceKeys.Contains(source)&&!levelKeys.Contains(key);
                };
                int removed=0;
                foreach(var level in settings.Levels.Where(l=>missing(l.Key)).ToList()){settings.Levels.Remove(level);removed++;}
                if(missing(settings.ZeroLevel)){settings.ZeroLevel=null;removed++;}
                if(missing(settings.GroundLevel)){settings.GroundLevel=null;removed++;}
                removed+=settings.Sources.RemoveAll(s=>!sourceKeys.Contains(s.Key));
                return removed;
            }
        }
    }
}
