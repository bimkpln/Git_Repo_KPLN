using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public class RoomCategoryWrite
        {
            public DepartmentRoom Room {get;set;}
            public string Category {get;set;}
            public string Source {get{return Room.Source;}}
            public string Element {get{return Room.Element;}}
            public string Number {get{return Room.Number;}}
            public string Name {get{return Room.Name;}}
            public string Level {get{return Room.Level;}}
        }
        public class RoomParameterWrite
        {
            public string SourceKey {get;set;}
            public string Source {get;set;}
            public string Element {get;set;}
            public string UniqueId {get;set;}
            public string Name {get;set;}
            public string Before {get;set;}
            public string After {get;set;}
            public string Status {get;set;}
            public bool Writable {get;set;}
        }
        public partial class Engine
        {
            // Единственное имя параметра, в который плагин записывает категории помещений.
            public const string DefaultRoomOutputParameter = "pg__ТЭП_НАЗНАЧЕНИЕ";
            public static string NormalizeRoomOutputParameter(string name)
            {return DefaultRoomOutputParameter;}

            public List<RoomCategoryWrite> RoomCategoryAssignments(IEnumerable<DepartmentGroup> groups,UserDictionary draft=null)
            {
                if(draft!=null)ValidateCategorySettings(draft);
                var result=new List<RoomCategoryWrite>();
                foreach(var group in groups)
                foreach(var room in group.Rooms)
                {
                    string role=DepartmentRole(group.Value);
                    if(draft!=null)
                    {
                        var parts=(group.Value??"").Split(new[]{" | "},2,StringSplitOptions.None);
                        var assignment=Config.Departments.FirstOrDefault(a=>Eq(a.Value,group.Value)&&(string.IsNullOrEmpty(a.Scope)||a.Scope==ClassificationScope));
                        var explicitRoles=assignment==null?new List<string>():draft.Categories.Where(c=>(parts.Length==2?c.Names:c.Departments)?.Any(v=>Eq(v,parts.Last()))==true).Select(c=>Roles().First(r=>Eq(r.Label,c.Category)).Key).Distinct().ToList();
                        string department=room.RawDepartment??parts[0];
                        role=explicitRoles.Count==1?explicitRoles[0]:assignment?.Role=="unknown"?"unknown":DictionaryRoomRole(draft,department,room.RawName??(parts.Length==2?parts[1]:room.Name),room.IsApartment||Eq(department,"Квартира"));
                        if(role=="unknown"&&assignment?.Role!="unknown"&&Eq(department,"Квартира"))role="heated";
                    }
                    bool families=draft==null?CategoryUsesFamilies(role):draft.Sources?.Any(s=>s.Role==role&&s.Families)==true;
                    if(role!="unknown"&&!families)result.Add(new RoomCategoryWrite{Room=room,Category=Roles().First(r=>r.Key==role).Label});
                }
                return result.OrderBy(r=>r.Category).ThenBy(r=>r.Source).ThenBy(r=>r.Level).ThenBy(r=>r.Number).ToList();
            }

            public static List<RoomCategoryWrite> NormalizeRoomCategoryWrites(IEnumerable<RoomCategoryWrite> assignments)
            {
                var result=new List<RoomCategoryWrite>();
                foreach(var group in assignments.GroupBy(r=>r.Room.SourceKey+"/"+(r.Room.UniqueId??r.Room.Element)))
                {
                    if(group.Any(r=>!Roles().Any(role=>role.Key!="unknown"&&role.Label==r.Category)))
                        throw new InvalidOperationException("Для записи нужна назначенная категория помещения.");
                    if(group.Select(r=>r.Category).Distinct().Count()!=1)
                        throw new InvalidOperationException("Помещение ID "+group.First().Element+" отнесено к нескольким категориям. Запись отменена.");
                    result.Add(group.First());
                }
                return result;
            }

            public List<RoomParameterWrite> PreviewRoomCategoryWrite(IEnumerable<RoomCategoryWrite> assignments)
            {
                return NormalizeRoomCategoryWrites(assignments).GroupBy(r=>r.Category)
                    .SelectMany(g=>PreviewRoomParameterWrite(g.Select(r=>r.Room),DefaultRoomOutputParameter,g.Key)).ToList();
            }
            public void WriteRoomCategories(IEnumerable<RoomCategoryWrite> assignments)
            {WriteRoomParameterRows(PreviewRoomCategoryWrite(assignments),DefaultRoomOutputParameter);}
            public static UserDictionary MergeCategoryDictionaries(UserDictionary shared,UserDictionary local,UserDictionary basis)
            {
                var merged=Deserialize<UserDictionary>(Serialize(shared));
                if(local==null)return merged;
                foreach(var category in local.Categories)
                {
                    var previous=basis?.Categories.FirstOrDefault(c=>Eq(c.Category,category.Category));
                    var target=merged.Categories.FirstOrDefault(c=>Eq(c.Category,category.Category));
                    if(target==null){target=new UserDictionaryCategory{Category=category.Category,Departments=new List<string>(),Names=new List<string>()};merged.Categories.Add(target);}
                    if(previous==null||category.ApartmentOnly!=previous.ApartmentOnly)target.ApartmentOnly=category.ApartmentOnly;
                    if(previous==null||Serialize(category.DisabledTags)!=Serialize(previous.DisabledTags))target.DisabledTags=category.DisabledTags?.ToList();
                    foreach(bool names in new[]{false,true})
                    {
                        Func<UserDictionaryCategory,List<string>> list=c=>names?c.Names:c.Departments;
                        var current=list(category)??new List<string>();var before=previous==null?new List<string>():list(previous)??new List<string>();
                        var result=list(target)??new List<string>();if(names)target.Names=result;else target.Departments=result;
                        foreach(var removed in before.Where(v=>!current.Any(x=>Eq(x,v))))result.RemoveAll(v=>Eq(v,removed));
                        foreach(var added in current.Where(v=>previous==null||!before.Any(x=>Eq(x,v))))
                        {
                            foreach(var other in merged.Categories.Where(c=>c!=target&&c.ApartmentOnly==target.ApartmentOnly))list(other)?.RemoveAll(v=>Eq(v,added));
                            if(!result.Any(v=>Eq(v,added)))result.Add(added);
                        }
                    }
                }
                if(basis==null||!Eq(local.OutputParameter,basis.OutputParameter))merged.OutputParameter=local.OutputParameter;
                merged.Sources=local.Sources;merged.Families=local.Families;
                return merged;
            }
            // Use the same authenticated department as the other KPLN modules. No local override.
            public static bool CanManageCategoryDictionary
            {
                get
                {
                    try
                    {
                        var type=AppDomain.CurrentDomain.GetAssemblies().Select(a=>a.GetType("KPLN_Loader.Application",false)).FirstOrDefault(t=>t!=null);
                        var department=type?.GetProperty("CurrentSubDepartment")?.GetValue(null)??type?.GetField("CurrentSubDepartment")?.GetValue(null);
                        if(department!=null)return Convert.ToInt32(department.GetType().GetProperty("Id")?.GetValue(department)??department.GetType().GetField("Id")?.GetValue(department))==8;
                        var service=AppDomain.CurrentDomain.GetAssemblies().Select(a=>a.GetType("KPLN_Library_DBWorker.SQLiteMainService",false)).FirstOrDefault(t=>t!=null);
                        department=service?.GetProperty("CurrentUserDBSubDepartment")?.GetValue(null)??service?.GetField("CurrentUserDBSubDepartment")?.GetValue(null);
                        return department!=null&&Convert.ToInt32(department.GetType().GetProperty("Id")?.GetValue(department)??department.GetType().GetField("Id")?.GetValue(department))==8;
                    }
                    catch{return false;}
                }
            }
            public UserDictionary CategorySettingsDraft()
            {
                var draft=userDictionary==null?new UserDictionary{Categories=new List<UserDictionaryCategory>()}:Deserialize<UserDictionary>(Serialize(userDictionary));
                foreach(var role in Roles().Where(r=>r.Key!="unknown"))
                    if(!draft.Categories.Any(c=>Eq(c.Category,role.Label)))draft.Categories.Add(new UserDictionaryCategory{Category=role.Label,Departments=new List<string>(),Names=new List<string>()});
                draft.OutputParameter=NormalizeRoomOutputParameter(draft.OutputParameter);
                draft.Sources=Deserialize<List<CategorySource>>(Serialize(Config.CategorySources));
                draft.Families=Deserialize<List<CategoryFamily>>(Serialize(Config.CategoryFamilies));
                // Include the current card assignments in the editable dictionary instead of hiding them behind it.
                foreach(var assignment in Config.Departments.Where(d=>d.Role!="unknown"&&!string.IsNullOrWhiteSpace(d.Value)&&(string.IsNullOrEmpty(d.Scope)||d.Scope==ClassificationScope)))
                {
                    var parts=assignment.Value.Split(new[]{" | "},2,StringSplitOptions.None);
                    var category=draft.Categories.First(c=>Eq(c.Category,Roles().First(r=>r.Key==assignment.Role).Label));
                    string value=parts.Last();bool name=parts.Length==2;
                    foreach(var other in draft.Categories)
                    {
                        if(name)other.Names?.RemoveAll(v=>Eq(v,value));else other.Departments?.RemoveAll(v=>Eq(v,value));
                    }
                    if(name){category.Names=category.Names??new List<string>();category.Names.Add(value);}
                    else{category.Departments=category.Departments??new List<string>();category.Departments.Add(value);}
                }
                return draft;
            }
            public static void ValidateCategorySettings(UserDictionary draft)
            {
                ValidateUserDictionary(draft);
                if(draft.Categories.GroupBy(c=>c.Category,StringComparer.OrdinalIgnoreCase).Any(g=>g.Count()>1))throw new InvalidOperationException("Категория повторяется в словаре.");
                ValidateOutputParameterName(NormalizeRoomOutputParameter(draft.OutputParameter));
                var roles=new HashSet<string>(Roles().Where(r=>r.Key!="unknown").Select(r=>r.Key));
                if((draft.Sources??new List<CategorySource>()).Any(c=>c==null||!roles.Contains(c.Role))||(draft.Sources??new List<CategorySource>()).GroupBy(c=>c.Role).Any(g=>g.Count()>1))throw new InvalidOperationException("Некорректные источники категорий.");
                var families=draft.Families??new List<CategoryFamily>();
                if(families.Any(f=>f==null||!roles.Contains(f.Role)||string.IsNullOrWhiteSpace(f.Category)||string.IsNullOrWhiteSpace(f.Family)||string.IsNullOrWhiteSpace(f.Type)||f.AreaFromParameter&&string.IsNullOrWhiteSpace(f.AreaParameter)))throw new InvalidOperationException("Заполните типы семейств и имена параметров площади.");
                if(families.GroupBy(f=>f.Caption,StringComparer.OrdinalIgnoreCase).Any(g=>g.Count()>1))throw new InvalidOperationException("Один тип семейства указан несколько раз. Оставьте одно назначение, чтобы исключить двойной учёт.");
            }
            public void ApplyCategorySettings(UserDictionary draft)
            {
                ValidateCategorySettings(draft);
                if(!CanManageCategoryDictionary&&!Eq(NormalizeRoomOutputParameter(draft.OutputParameter),NormalizeRoomOutputParameter(userDictionary?.OutputParameter))&&!Eq(draft.OutputParameter,SharedOutputParameterName()))throw new InvalidOperationException("Имя параметра записи в настройках меняет только BIM (8).");
                var copy=Deserialize<UserDictionary>(Serialize(draft));
                Config.LocalClassificationDictionary=copy;
                Config.CategorySources=copy.Sources??new List<CategorySource>();
                Config.CategoryFamilies=copy.Families??new List<CategoryFamily>();
                Config.CategorySettingsConfigured=true;
                foreach(var assignment in Config.Departments.Where(d=>string.IsNullOrEmpty(d.Scope)||d.Scope==ClassificationScope).ToList())
                {
                    var parts=(assignment.Value??"").Split(new[]{" | "},2,StringSplitOptions.None);
                    var matches=copy.Categories.Where(c=>(parts.Length==2?c.Names:c.Departments)?.Any(v=>Eq(v,parts.Last()))==true).Select(c=>Roles().First(r=>Eq(r.Label,c.Category)).Key).Distinct().ToList();
                    if(matches.Count==1)assignment.Role=matches[0];
                    else if(assignment.Role!="unknown")Config.Departments.Remove(assignment);
                }
                LoadClassificationDictionary();
            }
            public static string ReadDictionaryRevision()
            {return File.Exists(ClassificationDictionaryPath)?File.ReadAllText(ClassificationDictionaryPath):null;}
            public void PublishCategoryDictionary(UserDictionary draft,string expectedRevision)
            {
                if(!CanManageCategoryDictionary)throw new InvalidOperationException("Обновление общего словаря доступно только подразделению BIM (8).");
                ValidateCategorySettings(draft);
                string directory=Path.GetDirectoryName(ClassificationDictionaryPath);
                string temporary=Path.Combine(directory,"settings."+Guid.NewGuid().ToString("N")+".tmp");
                // The exclusive lock covers the compare and replace for all plugin writers.
                using(var guard=new FileStream(ClassificationDictionaryPath+".lock",FileMode.OpenOrCreate,FileAccess.ReadWrite,FileShare.None))
                {
                    if(!string.Equals(ReadDictionaryRevision(),expectedRevision,StringComparison.Ordinal))throw new InvalidOperationException("Общий словарь изменён другим пользователем. Откройте настройки заново: свежий словарь будет прочитан автоматически.");
                    try
                    {
                        using(var stream=File.Create(temporary))
                        using(var writer=JsonReaderWriterFactory.CreateJsonWriter(stream,Encoding.UTF8,false,true,"  "))
                            new DataContractJsonSerializer(typeof(UserDictionary)).WriteObject(writer,draft);
                        ValidateCategorySettings(Deserialize<UserDictionary>(File.ReadAllText(temporary)));
                        if(File.Exists(ClassificationDictionaryPath))File.Replace(temporary,ClassificationDictionaryPath,ClassificationDictionaryPath+"."+DateTime.Now.ToString("yyyyMMdd_HHmmss_fff")+".bak");
                        else File.Move(temporary,ClassificationDictionaryPath);
                    }
                    finally{if(File.Exists(temporary))File.Delete(temporary);}
                }
            }
            public static void ValidateOutputParameterName(string name)
            {
                if(string.IsNullOrWhiteSpace(name)||name.Trim()!=name||name.IndexOfAny(new[]{'\r','\n','\t'})>=0)throw new InvalidOperationException("Введите имя отдельного текстового параметра без переносов строк и пробелов по краям.");
                if(name.StartsWith("@")||new[]{"Назначение","ПОМ_Корпус","ПОМ_Секция","КВ_Номер","Имя","Номер"}.Any(n=>Eq(n,name)))throw new InvalidOperationException("Для записи выберите отдельный текстовый параметр. Исходное Назначение и идентификаторы помещений используются в расчёте.");
            }
            public List<string> RoomOutputParameters()
            {
                var names=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var iterator=doc.ParameterBindings.ForwardIterator();
                while(iterator.MoveNext())
                {
                    var definition=iterator.Key;var binding=iterator.Current as InstanceBinding;
                    if(binding==null||!binding.Categories.Contains(doc.Settings.Categories.get_Item(BuiltInCategory.OST_Rooms))||!IDHelper.IsText(definition))continue;
                    try{ValidateOutputParameterName(definition.Name);names.Add(definition.Name);}catch(InvalidOperationException){}
                }
                names.Add(NormalizeRoomOutputParameter(userDictionary?.OutputParameter));
                string sharedName=SharedOutputParameterName();if(!string.IsNullOrWhiteSpace(sharedName))names.Add(sharedName);
                return names.OrderBy(n=>n).ToList();
            }
            private static string SharedOutputParameterName()
            {
                try
                {
                    string text=ReadDictionaryRevision();if(text==null)return null;
                    var shared=Deserialize<UserDictionary>(text);ValidateCategorySettings(shared);
                    return NormalizeRoomOutputParameter(shared.OutputParameter);
                }
                catch{return null;}
            }
            public List<RoomParameterWrite> PreviewRoomParameterWrite(IEnumerable<DepartmentRoom> rooms,string name,string value)
            {
                ValidateOutputParameterName(name);
                if(string.IsNullOrWhiteSpace(value))throw new InvalidOperationException("Введите новое значение назначения.");
                if(!CanManageCategoryDictionary&&!RoomOutputParameters().Contains(name,StringComparer.OrdinalIgnoreCase))throw new InvalidOperationException("Выберите параметр из списка. Новое имя может задать только BIM (8).");
                var result=new List<RoomParameterWrite>();
                foreach(var item in rooms.GroupBy(r=>r.SourceKey+"/"+r.UniqueId).Select(g=>g.First()))
                {
                    var row=new RoomParameterWrite{SourceKey=item.SourceKey,Source=item.Source,Element=item.Element,UniqueId=item.UniqueId,Name=item.Name,After=value.Trim()};result.Add(row);
                    if(item.SourceKey!="host"){row.Status="Связь: запись недоступна";continue;}
                    try
                    {
                        var room=doc.GetElement(item.UniqueId) as Room;
                        if(room==null||!IsPlacedRoom(room)){row.Status="Помещение удалено или не размещено";continue;}
                        var parameter=InstanceParameter(room,name);
                        if(doc.IsReadOnly){row.Status="Документ открыт только для чтения";continue;}
                        if(parameter!=null&&(parameter.StorageType!=StorageType.String||parameter.IsReadOnly||!IDHelper.IsText(parameter.Definition))){row.Status="Параметр не текстовый или недоступен для записи";continue;}
                        row.Before=parameter?.AsString()??"";
                        row.Writable=true;row.Status=parameter==null?"Создать параметр и записать":row.Before==row.After?"Уже записано":"Записать новое значение";
                    }
                    catch(Exception ex){row.Status=ex.Message;}
                }
                return result;
            }
            public void WriteRoomParameter(IEnumerable<DepartmentRoom> selected,string name,string value)
            {
                WriteRoomParameterRows(PreviewRoomParameterWrite(selected,name,value),name);
            }
            private void WriteRoomParameterRows(IEnumerable<RoomParameterWrite> preview,string name)
            {
                var rows=preview.Where(r=>r.Writable).ToList();
                if(rows.Count==0)throw new InvalidOperationException("Нет доступных помещений основной модели для записи.");
                using(var transaction=new Transaction(doc,"ТЭП: запись назначения помещений"))
                {
                    transaction.Start();
                    var failures=new RoomWriteFailures();
                    transaction.SetFailureHandlingOptions(transaction.GetFailureHandlingOptions().SetFailuresPreprocessor(failures).SetClearAfterRollback(true));
                    EnsureRoomTextParameter(name,rows);
                    foreach(var row in rows)
                    {
                        var parameter=InstanceParameter(doc.GetElement(row.UniqueId),name);
                        if(parameter==null||parameter.IsReadOnly||parameter.StorageType!=StorageType.String)throw new InvalidOperationException("Не удалось записать параметр в помещение ID "+row.Element+". Изменения этой операции отменены.");
                        if(parameter.AsString()!=row.After&&!parameter.Set(row.After))throw new InvalidOperationException("Revit не принял значение у помещения ID "+row.Element+". Операция отменена.");
                    }
                    if(transaction.Commit()!=TransactionStatus.Committed)throw new InvalidOperationException("Revit отменил запись параметра. Изменения не применены. "+string.Join("; ",failures.Errors));
                }
            }
            private sealed class RoomWriteFailures : IFailuresPreprocessor
            {
                internal readonly List<string> Errors=new List<string>();
                public FailureProcessingResult PreprocessFailures(FailuresAccessor accessor)
                {
                    foreach(var failure in accessor.GetFailureMessages().Where(f=>f.GetSeverity()!=FailureSeverity.Warning))Errors.Add(failure.GetDescriptionText());
                    return Errors.Count>0?FailureProcessingResult.ProceedWithRollBack:FailureProcessingResult.Continue;
                }
            }
            private void EnsureRoomTextParameter(string name,List<RoomParameterWrite> rows)
            {
                if(rows.All(r=>InstanceParameter(doc.GetElement(r.UniqueId),name)!=null))return;
                var iterator=doc.ParameterBindings.ForwardIterator();
                while(iterator.MoveNext())if(Eq(iterator.Key.Name,name))throw new InvalidOperationException("Параметр с таким именем уже привязан, но недоступен у части помещений. Проверьте привязку параметра к экземплярам помещений.");
                string previous=doc.Application.SharedParametersFilename;
                string temporary=Path.Combine(Path.GetTempPath(),"TEP-parameter-"+Guid.NewGuid().ToString("N")+".txt");
                try
                {
                    File.WriteAllText(temporary,"");doc.Application.SharedParametersFilename=temporary;
                    var file=doc.Application.OpenSharedParameterFile();
                    if(file==null)throw new InvalidOperationException("Не удалось создать временный файл общего параметра.");
                    var group=file.Groups.Create("ТЭП");
                    using(var options=IDHelper.TextDefinitionOptions(name))
                    {
                        options.GUID=RoomOutputParameterGuid(name);
                        var definition=group.Definitions.Create(options);
                        var categories=doc.Application.Create.NewCategorySet();categories.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_Rooms));
                        if(!doc.ParameterBindings.Insert(definition,doc.Application.Create.NewInstanceBinding(categories)))throw new InvalidOperationException("Не удалось привязать текстовый параметр к помещениям.");
                    }
                    doc.Regenerate();
                }
                finally{doc.Application.SharedParametersFilename=previous;if(File.Exists(temporary))File.Delete(temporary);}
            }
            public static Guid RoomOutputParameterGuid(string name)
            {
                using(var hash=System.Security.Cryptography.SHA256.Create())
                    return new Guid(hash.ComputeHash(Encoding.UTF8.GetBytes("KPLN.CalculateTEP.RoomOutput/"+name.Trim().ToUpperInvariant())).Take(16).ToArray());
            }
        }
    }
}
