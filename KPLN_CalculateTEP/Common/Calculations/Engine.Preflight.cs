using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public partial class Engine
        {
            public static HashSet<string> RequiredParameterKeys(Settings settings,IEnumerable<string> objectRoles)
            {
                var roles=new HashSet<string>(objectRoles);var enabled=settings.Metrics.Where(m=>m.Enabled).Select(m=>m.Key).ToList();
                var required=new HashSet<string>();
                Action<string,bool> add=(key,needed)=>{if(needed&&!Settings.IsFixedParameter(key)&&enabled.Any(m=>ParameterAffectsMetric(key,m)))required.Add(key);};
                bool publicBuilding=OneOf(settings.Profile,"public","high-public");
                add("embedded",true);add("standalone",true);
                add("vertical",roles.Overlaps(new[]{"multilight","opening","shaft","engineering-shaft","stair-gap"}));
                add("width",roles.Overlaps(new[]{"arch","stair-gap"}));add("stair-width",roles.Contains("stair-gap"));
                bool technical=roles.Overlaps(new[]{"technical-space","technical-void"});
                add("height",roles.Contains("niche")||publicBuilding&&technical&&enabled.Any(m=>GrossMetric((Indicator)Enum.Parse(typeof(Indicator),m))));
                add("roof-ratio",publicBuilding&&roles.Contains("roof-vent"));
                add("service-access",publicBuilding&&technical);
                add("mezzanine-ratio",publicBuilding&&roles.Contains("mezzanine"));
                add("transition",!publicBuilding&&roles.Contains("transition"));
                return required;
            }
            public string RoomScanSummary { get; private set; }
            public bool HasSummerRooms { get; private set; }

            private List<Record> PreflightRecords()
            {
                Progress("Предварительная проверка: параметры, помещения и автоматические размеры...");
                RoomScanSummary = ""; HasSummerRooms = false;
                if(singleBuildingAssumption!=null)
                    current.Issue("SINGLE_BUILDING_ASSUMPTION","Предупреждение",
                        "Пользователь подтвердил расчёт выбранных источников как одного корпуса «"+singleBuildingAssumption.BuildingName+"». "+singleBuildingAssumption.Description+" Корпус назначен только внутри расчёта; параметры помещений не изменены.",
                        action:"Результат получен при допущении единого корпуса. Допущение действует только для этой проверки или расчёта и не сохраняется в настройках.");
                if (!Config.Metrics.Any(m => m.Enabled)) current.Issue("NO_METRICS", "Ошибка", "Не выбраны показатели расчёта.");
                LoadClassificationDictionary();
                if(DictionaryStatus.StartsWith("Словарь недоступен"))current.Issue("CLASSIFICATION_DICTIONARY","Предупреждение",DictionaryStatus);
                var records = Collect();
                automaticGeometryRecords=records;automaticDimensions.Clear();automaticDimensionErrors.Clear();
                foreach(var category in (Config.CategorySources??new List<CategorySource>()).Where(c=>c.Families))
                    if(!records.Any(r=>IsClassifiedFamily(r.Element)&&r.Role==category.Role))
                        current.Issue("CATEGORY_FAMILY_EMPTY","Ошибка","Для категории «"+(Roles().FirstOrDefault(r=>r.Key==category.Role)?.Label??category.Role)+"» выбран расчёт через семейства, но нет пригодных экземпляров выбранных типов. Категория пропущена; остальные данные рассчитываются.");
                foreach (var level in Config.Levels.Where(l => l.Include))
                    foreach (var entry in new[] { Tuple.Create("Верх перекрытия", level.TopSlabMode=="manual"?level.TopSlab:null), Tuple.Create("Высота", level.HeightMode=="manual"?level.Height:null), Tuple.Create("Доля от кровли", level.RoofRatioMode=="manual"?level.RoofRatio:null), Tuple.Create("Площадь надстройки", level.RoofAreaMode=="manual"?level.RoofArea:null) })
                    {
                        if (string.IsNullOrWhiteSpace(entry.Item2)) continue;
                        double value;
                        if (!TryNumber(entry.Item2, out value) || entry.Item1 != "Верх перекрытия" && value < 0 || entry.Item1 == "Доля от кровли" && value > 1)
                            current.Issues.Add(new Issue{Code="LEVEL_VALUE",Severity="Ошибка",Message="Этаж «"+level.Name+"»: недопустимое значение поля «"+entry.Item1+"».",Source=level.Source,ParameterKey=entry.Item1=="Верх перекрытия"?"$level-top":"$level-normative"});
                    }
                foreach (var r in records.Where(r => IsAreaInput(r.Element) && r.Level == null))
                    Notice("ROOM_LEVEL", "Ошибка", "Не определён расчётный уровень помещения.", r);
                var rooms = records.Where(r => IsAreaInput(r.Element) && r.Level != null && r.Level.Include && r.Level.Kind != "exclude").ToList();
                if (rooms.Count == 0)
                {
                    var reasons=new Dictionary<string,int>(roomExclusionReasons);
                    int noLevel=records.Count(r=>IsAreaInput(r.Element)&&r.Level==null);
                    int excludedLevel=records.Count(r=>IsAreaInput(r.Element)&&r.Level!=null&&(!r.Level.Include||r.Level.Kind=="exclude"));
                    if(noLevel>0)reasons["У объекта не определён расчётный уровень."]=noLevel;
                    if(excludedLevel>0)reasons["Этаж объекта выключен из расчёта или помечен как исключённый."]=excludedLevel;
                    current.Issue("ROOMS_EMPTY", "Ошибка", EmptyAreaInputsMessage(placedRoomsScanned,reasons));
                }
                foreach (var group in records.GroupBy(r => r.Building))
                {
                    bool single = group.Where(r => r.AutomaticShaftRegion == null && r.Level != null && r.Level.Include && r.Role != "mezzanine").Select(r => Math.Round(r.Z, 6)).Distinct().Count() == 1;
                    foreach (var r in group) r.SingleStorey = single;
                }
                bool apartmentMetrics = Config.Metrics.Any(m => m.Enabled && m.Key.StartsWith("Apartments"));
                if (apartmentMetrics && string.IsNullOrWhiteSpace(Config.Parameter("apartment")))
                    current.Issue("APARTMENT_PARAMETER", "Ошибка", "Назначьте параметр сквозного номера / ID квартиры. Один ID объединяет этажи одной квартиры; у разных квартир ID должны различаться в пределах корпуса и секции.");
                foreach (var source in Sources.Where(s => s.Loaded && s.LoadError == null))
                {
                    var sourceRooms = rooms.Where(r => r.Source == source).ToList();
                    if (sourceRooms.Count == 0) continue;
                    if (sourceRooms.Any(r=>r.Element is Room) && Config.Metrics.Any(m => m.Enabled && RoomAreaMetric((Indicator)Enum.Parse(typeof(Indicator), m.Key))) &&
                        AreaVolumeSettings.GetAreaVolumeSettings(source.Document).GetSpatialElementBoundaryLocation(SpatialElementType.Room) != SpatialElementBoundaryLocation.Finish)
                        current.Issue("ROOM_AREA_SETTINGS", "Ошибка", "Площадь Room рассчитывается не по чистовой грани. Для использования Room.Area требуется способ вычисления границ помещений «По отделке стен». Настройки модели автоматически не изменяются.", source: source.Name);
                }

                if (Config.Metrics.Any(m => m.Enabled && m.Key == "ParkingCount"))
                    foreach(var parking in records.Where(r => IsParkingFamily(r.Element)))
                    {
                        if(parking.Level==null)Notice("PARKING_LEVEL", "Ошибка", "У семейства категории «Парковка» не определён уровень.", parking, "ParkingCount");
                        if(parking.Level!=null && (!parking.Level.Include || parking.Level.Kind=="exclude"))continue;
                        if(string.IsNullOrWhiteSpace(Mapped(parking.Element,"parking")))Notice("PARKING_MARK", "Ошибка", "У семейства категории «Парковка» не заполнена «Марка» экземпляра - порядковый номер машино-места.", parking, "ParkingCount");
                    }
                int index = 0;
                var loggias = new Dictionary<string, double>(); var balconies = new Dictionary<string, double>();
                foreach (var original in rooms)
                {
                    Progress("Проверка заполнения помещений: " + (++index) + " / " + rooms.Count);
                    foreach(var key in new[]{"embedded","standalone","service-access"})
                    {
                        if(!Config.Metrics.Any(m=>m.Enabled&&ParameterAffectsMetric(key,m.Key)))continue;
                        try
                        {
                            if(string.IsNullOrWhiteSpace(Mapped(original.Element,key)))continue;
                            if(OneOf(key,"include","embedded","standalone","partial-floor","service-access")){Flag(original.Element,key);continue;}
                            bool length=OneOf(key,"height","width","stair-width","floor-offset");
                            var n=Number(original.Element,key,length);
                            if(!n.HasValue||double.IsNaN(n.Value)||double.IsInfinity(n.Value)||key!="floor-offset"&&n<0||OneOf(key,"roof-ratio","mezzanine-ratio")&&n>1||key=="slope"&&n>90)
                                throw new InvalidOperationException("Недопустимое значение параметра «"+Config.Parameter(key)+"».");
                        }
                        catch(OperationCanceledException){throw;}
                        catch(Exception ex){ParameterIssue("PARAM_VALUE",ex.Message,key,original);}
                    }
                    foreach (var metric in Config.Metrics.Where(m => m.Enabled))
                    {
                        if(metric.Key=="ParkingCount")continue;
                        var r = original.Copy();
                        try
                        {
                            ApplyRules(r, metric.Key);
                            var indicator = (Indicator)Enum.Parse(typeof(Indicator), metric.Key);
                            if (r.Role == "unknown" && metric.Key!="ApartmentsCount") throw new InvalidOperationException("Назначение помещения не распределено. Откройте карточку «Нераспределённые» в классификации и назначьте категорию.");
                            if (metric.Key.StartsWith("Apartments") && OneOf(r.Role, "heated", "auxiliary") && string.IsNullOrWhiteSpace(r.Apartment))
                                throw new InvalidOperationException("У квартирного помещения отсутствует сквозной номер / ID квартиры.");
                            if(metric.Key.StartsWith("Apartments") && !string.IsNullOrWhiteSpace(r.Apartment) && string.IsNullOrWhiteSpace(r.Section))
                                throw new InvalidOperationException("У квартиры не заполнена секция «ПОМ_Секция».");
                            if(metric.Key=="ApartmentsTotal" && !string.IsNullOrWhiteSpace(r.Apartment))Factor(r,indicator);
                            HasSummerRooms |= OneOf(r.Role, "loggia", "balcony", "terrace");
                            if (metric.Key == "ApartmentsTotal" && !string.IsNullOrWhiteSpace(r.Apartment))
                            {
                                if (r.Role == "loggia") loggias[r.Key] = (r.Element is Room?((Room)r.Element).Area:UsesFamilyParameterArea(r.Element)?FamilyParameterArea(r):0) * .09290304;
                                if (OneOf(r.Role, "balcony", "terrace")) balconies[r.Key] = (r.Element is Room?((Room)r.Element).Area:UsesFamilyParameterArea(r.Element)?FamilyParameterArea(r):0) * .09290304;
                            }
                            if(!CountMetric(metric.Key))
                            {
                            if (OneOf(r.Role, "multilight", "opening", "shaft", "engineering-shaft", "stair-gap") && string.IsNullOrWhiteSpace(r.Vertical))
                                throw new InvalidOperationException("Не заполнен общий ID вертикального пространства.");
                            if (r.Role == "stair-gap") { Dimension(r, "width", true); Dimension(r, "stair-width", true); }
                            if (r.Role == "niche") RoomClearance(r);
                            if (r.Role == "arch") Dimension(r, "width", true);
                            if (NeedsCeilingCheck(r,indicator) && HasSlopedCeiling(r) && !IsPublic(r) && r.Role != "public" && r.Part != "nonresidential")
                                throw new InvalidOperationException("Для жилой мансарды исходное ТЗ неоднозначно задаёт высотные пороги. Нужен согласованный порядок учёта; площадь автоматически не подменяется.");
                            }
                            if (OneOf(metric.Key, "Storeys", "Floors")) CountLevel(r, metric.Key == "Storeys");
                            else if (!metric.Key.StartsWith("Volume") && !OneOf(metric.Key, "Footprint", "ApartmentsCount", "ParkingCount")) Eligible(r, indicator);
                        }
                        catch (OperationCanceledException) { throw; }
                        catch (Exception ex) { Notice("PREFLIGHT_RULE", "Ошибка", ex.Message, r, metric.Key); }
                    }
                }
                var flats = rooms.Where(r => !string.IsNullOrWhiteSpace(r.Apartment)).GroupBy(ApartmentKey).ToList();
                int multilevel = flats.Count(g => g.Select(r => Math.Round(r.Z, 6)).Distinct().Count() > 1);
                RoomScanSummary = "Проверено помещений: " + rooms.Count + ". Квартир: " + flats.Count + ". Квартир на нескольких этажах: " + multilevel + ".";
                RoomScanSummary += "\nЛоджии квартир: " + loggias.Count + " (" + loggias.Values.Sum().ToString("0.###") + " м² по Revit). Балконы и террасы квартир: " + balconies.Count + " (" + balconies.Values.Sum().ToString("0.###") + " м² по Revit). Площади до коэффициентов и высотных исключений.";
                current.Issue("ROOM_SCAN", "Информация", RoomScanSummary);
                if (multilevel > 0) current.Issue("APARTMENT_MULTILEVEL", "Предупреждение", "Одинаковые номера в одном корпусе и секции объединены между этажами. Убедитесь, что это одна многоэтажная квартира, а не повторная нумерация разных квартир на этажах.");
                return records;
            }

            // Source is retained only when the building is inferred from the source. An explicit
            // building + section + apartment ID also groups apartments split between linked models.
            private string ApartmentKey(Record r)
            {
                return ApartmentIdentity(string.IsNullOrWhiteSpace(Config.Parameter("building")) ? r.Source.Key : "", r.Building, r.Section, r.Apartment);
            }
            public static string ApartmentIdentity(string source, string building, string section, string apartment)
            {
                return string.Concat(new[] { source, building, section, apartment }.Select(s => { var v = (s ?? "").Trim().ToUpperInvariant(); return v.Length + ":" + v; }));
            }
            public static double SummerCoefficient(string text)
            {
                double value;
                if (!TryNumber(text, out value) || value < 0 || value > 1) throw new InvalidOperationException("Коэффициент летнего помещения должен быть числом от 0 до 1.");
                return value;
            }
        }
    }
}
