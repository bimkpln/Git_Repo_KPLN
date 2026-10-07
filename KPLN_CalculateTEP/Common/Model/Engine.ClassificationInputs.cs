using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public class ClassificationDictionary
        {
            public int Version { get; set; } = 1;
            public List<DepartmentAssignment> Entries { get; set; } = new List<DepartmentAssignment>();
        }
        public partial class Engine
        {
            public const string ClassificationDictionaryPath = @"Z:\Отдел BIM\07_Вспомогательное\ТЭП_Plugin\settings.json";
            private ClassificationDictionary classificationDictionary = new ClassificationDictionary();
            public string DictionaryStatus { get; private set; } = "Словарь ещё не прочитан.";
            public string DictionaryRevision {get;private set;}
            public void LoadClassificationDictionary()
            {
                try
                {
                    string revision=File.ReadAllText(ClassificationDictionaryPath);
                    var shared=Deserialize<UserDictionary>(revision);ValidateCategorySettings(shared);
                    var dictionary=MergeCategoryDictionaries(shared,Config.LocalClassificationDictionary,Config.SharedClassificationSnapshot);
                    ValidateCategorySettings(dictionary);
                    DictionaryRevision=revision;
                    if(Config.LocalClassificationDictionary!=null)Config.LocalClassificationDictionary=Deserialize<UserDictionary>(Serialize(dictionary));
                    Config.SharedClassificationSnapshot=Deserialize<UserDictionary>(Serialize(shared));
                    if(!Config.CategorySettingsConfigured)
                    {
                        if(Config.CategorySources.Count==0&&Config.CategoryFamilies.Count==0)
                        {
                            Config.CategorySources=Deserialize<List<CategorySource>>(Serialize(dictionary.Sources??new List<CategorySource>()));
                            Config.CategoryFamilies=Deserialize<List<CategoryFamily>>(Serialize(dictionary.Families??new List<CategoryFamily>()));
                        }
                        Config.CategorySettingsConfigured=true;
                    }
                    userDictionary=dictionary;
                    classificationDictionary=new ClassificationDictionary{Entries=dictionary.Categories.Where(c=>!c.ApartmentOnly).SelectMany(c=>(c.Departments??new List<string>()).Select(v=>new DepartmentAssignment{Value=v,Role=Roles().First(r=>Eq(r.Label,c.Category)).Key})).ToList()};
                    DictionaryStatus = "settings.json: "+dictionary.Categories.Count+" категорий; "+dictionary.Categories.Sum(c=>(c.Departments?.Count??0)+(c.Names?.Count??0))+" назначений и имён. "+ClassificationDictionaryPath;
                }
                catch (Exception ex)
                {
                    var fallback=Config.LocalClassificationDictionary??Config.SharedClassificationSnapshot;
                    if(fallback!=null)
                    {
                        ValidateCategorySettings(fallback);userDictionary=fallback;
                        classificationDictionary=new ClassificationDictionary{Entries=fallback.Categories.Where(c=>!c.ApartmentOnly).SelectMany(c=>(c.Departments??new List<string>()).Select(v=>new DepartmentAssignment{Value=v,Role=Roles().First(r=>Eq(r.Label,c.Category)).Key})).ToList()};
                        DictionaryStatus="Словарь недоступен: "+ex.Message+" Используется последняя сохранённая версия.";return;
                    }
                    classificationDictionary = new ClassificationDictionary();userDictionary=null;DictionaryRevision=null;
                    DictionaryStatus = "Словарь недоступен: " + ex.Message + " Используются ручные назначения и точные названия категорий.";
                }
            }
            public static void ValidateClassificationDictionary(ClassificationDictionary dictionary)
            {
                if (dictionary == null || dictionary.Version != 1 || dictionary.Entries == null)
                    throw new InvalidOperationException("Не поддерживается версия или структура словаря.");
                var roles = new HashSet<string>(Roles().Select(r => r.Key));
                if (dictionary.Entries.Any(e => e == null || string.IsNullOrWhiteSpace(e.Value) || !roles.Contains(e.Role)))
                    throw new InvalidOperationException("В словаре пустое значение или неизвестная категория Role.");
                if (dictionary.Entries.Any(e => !string.IsNullOrEmpty(e.Scope) && e.Scope!="apartment"))
                    throw new InvalidOperationException("Scope словаря: пусто для всех объектов либо apartment для комнат внутри назначения Квартира.");
                if (dictionary.Entries.GroupBy(e => e.Value.Trim(), StringComparer.OrdinalIgnoreCase).Any(g => g.Count() > 1))
                    throw new InvalidOperationException("В словаре повторяются значения (сравнение без учёта регистра).");
            }
            private static Parameter InstanceParameter(Element element, string name)
            {
                if (string.IsNullOrWhiteSpace(name)) return null;
                if (name == "@Department") return element.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT);
                if (name == "@Name" && element is Room) return element.get_Parameter(BuiltInParameter.ROOM_NAME);
                var parameters = element.GetParameters(name);
                if (parameters.Count > 1) throw new InvalidOperationException("Несколько параметров экземпляра с именем «" + name + "». Укажите однозначное имя.");
                return parameters.FirstOrDefault();
            }
            private static string InstanceText(Element element, string name)
            {
                var parameter = InstanceParameter(element, name);
                if (parameter == null) throw new InvalidOperationException("У экземпляра отсутствует параметр «" + name + "».");
                return (parameter.StorageType == StorageType.String ? parameter.AsString() : parameter.AsValueString())?.Trim() ?? "";
            }
            private bool IsClassifiedFamily(Element element)
            {
                return ClassifiedFamilyRole(element)!=null;
            }
            private bool IsAreaInput(Element element) { return element is Room || IsClassifiedFamily(element); }
            private string ClassificationValue(Element element)
            {
                if (element is Room)
                {
                    string value = InstanceText(element, "@Department");
                    if(Config.Departments?.Any(d=>Eq(d.Value,value)&&(string.IsNullOrEmpty(d.Scope)||d.Scope==ClassificationScope))==true)return value;
                    string name=InstanceText(element,"@Name");
                    string apartmentId=Mapped(element,"apartment");
                    bool apartment=Eq(Value(element,"@Department"),"Квартира")||!string.IsNullOrWhiteSpace(apartmentId)&&apartmentId!="0"&&apartmentId!="-";
                    if(DictionaryRoomRole("",name,apartment)!="unknown")return (apartment?"Квартира":value)+" | "+name;
                    return value;
                }
                return "Семейство | " + ClassifiedFamilyRole(element);
            }
            private string ClassificationScope { get { return "@Department|" + (Config.FamilyClassificationParameter ?? ""); } }
            private double FamilyParameterArea(Record record)
            {
                string name=FamilyAreaParameterName(record.Element);
                var p = InstanceParameter(record.Element, name);
                if (p == null || !p.HasValue) throw new InvalidOperationException("Не заполнен параметр площади экземпляра «" + name + "».");
                if (p.StorageType != StorageType.Double || !IDHelper.IsArea(p))
                    throw new InvalidOperationException("Параметр «" + name + "» должен иметь тип данных Площадь. Текст и безразмерные числа не пересчитываются в м² автоматически.");
                double area = p.AsDouble();
                if (double.IsNaN(area) || double.IsInfinity(area) || area < 0) throw new InvalidOperationException("Недопустимая площадь экземпляра.");
                return area;
            }
            private PlanarRegion FamilyAreaRegion(Record record)
            {
                if(UsesFamilyParameterArea(record.Element))throw new InvalidOperationException("Для экземпляра выбран параметр площади «"+FamilyAreaParameterName(record.Element)+"». Он подходит для суммирования площадей помещений, но не задаёт контур для геометрических проверок. Выберите геометрию семейства для этого показателя.");
                PlanarRegion edited;if(TryEditedObjectOverlay(record,out edited))return edited;
                double elevation = RoomFloorElevation(record);
                return ReviewRegion("family/" + record.Key, record, "Геометрия экземпляра", elevation, () =>
                {
                    var region = PlanarRegion.Empty;
                    foreach (var solid in Solids(record.Element))
                        region = RegionUnion(region, ProjectionRegion(SolidUtils.CreateTransformed(solid, record.Source.Transform)));
                    if (region.IsEmpty) throw new InvalidOperationException("У экземпляра нет площадной геометрии. Привяжите цветовую область в визуальном контроле или задайте параметр площади для суммирования.");
                    return region;
                });
            }
        }
    }
}
