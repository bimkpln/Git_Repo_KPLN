using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public class CategorySource
        {
            public string Role {get;set;}
            public bool Families {get;set;}
        }
        public class CategoryFamily
        {
            public string Role {get;set;}
            public string Category {get;set;}
            public string Family {get;set;}
            public string Type {get;set;}
            public bool AreaFromParameter {get;set;}
            public string AreaParameter {get;set;} = "Площадь";
            public string Caption {get{return Category+" / "+Family+" / "+Type;}}
        }
        public class FamilyChoice : CategoryFamily
        {
            public List<DepartmentRoom> Instances {get;set;} = new List<DepartmentRoom>();
            public string Display {get{return Caption+" ("+Instances.Count+")";}}
        }
        [DataContract]
        public class UserDictionary
        {
            [DataMember(Name="Версия",Order=0)] public int Version {get;set;} = 2;
            [DataMember(Name="Описание",Order=1)] public string Description {get;set;}
            [DataMember(Name="Категории",Order=2)] public List<UserDictionaryCategory> Categories {get;set;}
            [DataMember(Name="Параметр записи назначения",Order=3)] public string OutputParameter {get;set;} = Engine.DefaultRoomOutputParameter;
            [DataMember(Name="Источники категорий",Order=4)] public List<CategorySource> Sources {get;set;}
            [DataMember(Name="Семейства",Order=5)] public List<CategoryFamily> Families {get;set;}
        }
        [DataContract]
        public class UserDictionaryCategory
        {
            [DataMember(Name="Категория",Order=0)] public string Category {get;set;}
            [DataMember(Name="Назначения",Order=1)] public List<string> Departments {get;set;}
            [DataMember(Name="Имена помещений",Order=2)] public List<string> Names {get;set;}
            [DataMember(Name="Только внутри квартиры",Order=3)] public bool ApartmentOnly {get;set;}
            [DataMember(Name="Неактивные значения",Order=4,EmitDefaultValue=false)] public List<string> DisabledTags {get;set;}
        }
        public partial class Engine
        {
            private UserDictionary userDictionary;
            public List<FamilyChoice> FamilyChoices {get;private set;} = new List<FamilyChoice>();
            public bool CategoryUsesFamilies(string role){return Config.CategorySources?.Any(c=>c.Role==role&&c.Families)==true;}
            public void SetCategorySource(string role,bool families)
            {
                if(role=="unknown")return;
                Config.CategorySources=Config.CategorySources??new List<CategorySource>();
                Config.CategorySources.RemoveAll(c=>c.Role==role);
                Config.CategorySources.Add(new CategorySource{Role=role,Families=families});
            }
            public void AddCategoryFamily(string role,FamilyChoice choice)
            {
                if(role=="unknown"||choice==null)throw new InvalidOperationException("Выберите категорию и семейство.");
                Config.CategoryFamilies=Config.CategoryFamilies??new List<CategoryFamily>();
                if(Config.CategoryFamilies.Any(c=>SameFamily(c,choice)&&c.Role!=role))
                    throw new InvalidOperationException("Этот тип уже назначен другой категории. Удалите прежнее назначение, чтобы избежать двойного учёта.");
                if(!Config.CategoryFamilies.Any(c=>SameFamily(c,choice)))
                    Config.CategoryFamilies.Add(new CategoryFamily{Role=role,Category=choice.Category,Family=choice.Family,Type=choice.Type});
                SetCategorySource(role,true);
            }
            private CategoryFamily FamilyAreaSetting(Element element)
            {
                var instance=element as FamilyInstance;if(instance==null)return null;
                return Config.CategoryFamilies?.FirstOrDefault(c=>CategoryUsesFamilies(c.Role)&&Eq(c.Category,instance.Category?.Name)&&Eq(c.Family,instance.Symbol.FamilyName)&&Eq(c.Type,instance.Symbol.Name));
            }
            private bool UsesFamilyParameterArea(Element element){return FamilyAreaSetting(element)?.AreaFromParameter==true;}
            private string FamilyAreaParameterName(Element element){return FamilyAreaSetting(element)?.AreaParameter;}
            private static bool SameFamily(CategoryFamily a,CategoryFamily b)
            {return Eq(a.Category,b.Category)&&Eq(a.Family,b.Family)&&Eq(a.Type,b.Type);}
            private string ClassifiedFamilyRole(Element element)
            {
                var family=element as FamilyInstance;if(family==null)return null;
                var roles=(Config.CategoryFamilies??new List<CategoryFamily>()).Where(c=>CategoryUsesFamilies(c.Role)&&
                    Eq(c.Category,family.Category?.Name)&&Eq(c.Family,family.Symbol.FamilyName)&&Eq(c.Type,family.Symbol.Name)).Select(c=>c.Role).Distinct().ToList();
                if(roles.Count>1)throw new InvalidOperationException("Тип семейства назначен нескольким категориям.");
                return roles.FirstOrDefault();
            }
            public void ScanFamilyChoices(Action<string> progress=null)
            {
                var choices=new Dictionary<string,FamilyChoice>();int n=0;
                foreach(var source in Sources.Where(s=>s.Loaded&&s.LoadError==null&&s.Mode!="exclude"))
                    foreach(var instance in source.Elements.OfType<FamilyInstance>())
                    {
                        if(++n%100==0)progress?.Invoke("Список типов семейств: "+n);
                        if(!PhaseAccepted(source,instance))continue;
                        string category=instance.Category?.Name??"",family=instance.Symbol.FamilyName,type=instance.Symbol.Name;
                        string key=category+"\u001f"+family+"\u001f"+type;
                        FamilyChoice choice;
                        if(!choices.TryGetValue(key,out choice))choices[key]=choice=new FamilyChoice{Category=category,Family=family,Type=type};
                        choice.Instances.Add(new DepartmentRoom{SourceKey=source.Key,UniqueId=instance.UniqueId,Source=source.Name,Element=IDHelper.ElIdValue(instance.Id).ToString(),Name=family+" / "+type,
                            Level=(instance.Document.GetElement(instance.LevelId) as Level)?.Name,Department="Геометрия семейства"});
                    }
                FamilyChoices=choices.Values.OrderBy(c=>c.Caption).ToList();
            }
            public List<DepartmentRoom> CategoryFamilyInstances(string role)
            {
                return FamilyChoices.Where(c=>(Config.CategoryFamilies??new List<CategoryFamily>()).Any(a=>a.Role==role&&SameFamily(a,c))).SelectMany(c=>c.Instances).ToList();
            }
            public static void ValidateUserDictionary(UserDictionary dictionary)
            {
                if(dictionary==null||dictionary.Version!=2||dictionary.Categories==null)throw new InvalidOperationException("settings.json: ожидается Версия 2 и список Категории.");
                var labels=Roles().ToDictionary(r=>r.Label,r=>r.Key,StringComparer.OrdinalIgnoreCase);
                var used=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
                foreach(var category in dictionary.Categories)
                {
                    if(category==null||!labels.ContainsKey(category.Category??""))throw new InvalidOperationException("Неизвестное название категории в settings.json: "+category?.Category);
                    foreach(var group in new[]{Tuple.Create("назначение",category.Departments),Tuple.Create("имя",category.Names)})
                        foreach(var value in group.Item2??new List<string>())
                        {
                            if(string.IsNullOrWhiteSpace(value))throw new InvalidOperationException("Пустое значение в категории "+category.Category);
                            string key=group.Item1+"|"+category.ApartmentOnly+"|"+value.Trim(),previous;
                            if(used.TryGetValue(key,out previous))throw new InvalidOperationException("Повторяющееся значение «"+value+"»: "+previous+" / "+category.Category);
                            used.Add(key,category.Category);
                        }
                }
            }
            public string DictionaryRoomRole(string department,string name,bool apartment)
            {return DictionaryRoomRole(userDictionary,department,name,apartment);}
            private static string DictionaryRoomRole(UserDictionary dictionary,string department,string name,bool apartment)
            {
                if(dictionary==null)return "unknown";
                Func<UserDictionaryCategory,bool> accepted=c=>!c.ApartmentOnly||apartment;
                // Specific room names refine generic departments; manual card assignments are checked first.
                var matches=dictionary.Categories.Where(accepted).Where(c=>(c.Names??new List<string>()).Any(v=>Eq(v,name))).ToList();
                if(matches.Count==0)matches=dictionary.Categories.Where(accepted).Where(c=>(c.Departments??new List<string>()).Any(v=>Eq(v,department))).ToList();
                var roles=matches.Select(c=>Roles().First(r=>Eq(r.Label,c.Category)).Key).Distinct().ToList();
                return roles.Count==1?roles[0]:"unknown";
            }
        }
    }
}
