using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using KPLN_Tools.Common;
using KPLN_Tools.Forms;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Xml;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class Command_AR_calculateTEP : IExternalCommand
    {
        private static readonly Guid StorageSchemaGuid = new Guid("8C7E676D-14AB-47AE-B252-C0EC5C4F7D46");
        private const string StorageName = "KPLN_TEP_STORAGE";
        private const string SettingsField = "SettingsJson";
        private const string ResultField = "ResultJson";

        private sealed class SourceContext
        {
            public TepSourceOption Option;
            public Document Document;
            public Transform Transform;
        }

        private sealed class SpatialRecord
        {
            public string SourceKey, Source, Building, Level, Category, ElementName, Function, Apartment, BuildingType;
            public long ElementId, HostLinkId;
            public double Area, Volume, Elevation, Height, SummerCoefficient;
            public bool Include, Residential, Public, Technical, Summer, Underground, BuiltIn, Standalone, ApproximateVolume;
        }

        private sealed class ParkingRecord
        {
            public string SourceKey, Source, Building, Level, Name;
            public long ElementId, HostLinkId;
            public bool Underground;
        }

        private sealed class MetricValue
        {
            public double Value;
            public string Status = "Рассчитано";
            public string Comment = string.Empty;
        }

        private static readonly string[,] Indicators =
        {
            {"1","Суммарная поэтажная площадь всего","м²"},{"2","Суммарная поэтажная площадь жилых зданий","м²"},{"2а","Жилая часть суммарной поэтажной площади жилых зданий","м²"},{"2б","Нежилая часть суммарной поэтажной площади жилых зданий","м²"},{"3","Суммарная поэтажная площадь нежилых зданий","м²"},
            {"4","Площадь застройки","м²"},{"5","Строительный объём","м³"},{"5а","Строительный объём выше отметки 0.000","м³"},{"5б","Строительный объём ниже отметки 0.000","м³"},{"6","Общая площадь объекта","м²"},{"6а","Общая площадь наземной части","м²"},{"6б","Общая площадь подземной части","м²"},
            {"7","ННП встроенно-пристроенная","м²"},{"8","ННП встроенно-пристроенная отдельно стоящая","м²"},{"9","Наземная площадь","м²"},{"9а","Наземная площадь жилых зданий","м²"},{"9б","Наземная площадь нежилых зданий","м²"},{"10","Расчётная площадь общественного здания","м²"},
            {"11","Общая площадь квартир с летними помещениями","м²"},{"12","Площадь квартир без летних помещений","м²"},{"13","Площадь помещений общественного назначения","м²"},{"14","Количество квартир","шт."},{"15","Количество машино-мест в подземном паркинге","шт."},{"16","Этажность","этажей"},{"17","Количество этажей","этажей"}
        };

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc == null ? null : uidoc.Document;
            if (doc == null)
            {
                message = "Не найден активный документ Revit.";
                return Result.Failed;
            }

            try
            {
                TepSettings settings = LoadSettings(doc) ?? new TepSettings();
                while (true)
                {
                    List<SourceContext> sources = ScanSources(doc, settings);
                    AR_calculateTEP dialog = new AR_calculateTEP(uidoc, sources.Select(x => x.Option), settings);
                    bool? dialogResult = dialog.ShowDialog();
                    if (dialogResult != true || dialog.RequestedAction == TepDialogAction.None) return Result.Cancelled;

                    if (dialog.RequestedAction == TepDialogAction.ShowLast)
                    {
                        TepCalculationOutput last = LoadResult(doc);
                        if (last == null) TaskDialog.Show("Расчёт ТЭП", "В проекте ещё нет сохранённого расчёта.");
                        else ShowResultDialog(uidoc, last);
                        continue;
                    }
                    if (dialog.RequestedAction == TepDialogAction.ImportSettings)
                    {
                        settings = ReadJson<TepSettings>(dialog.RequestedPath);
                        SaveSettings(doc, settings);
                        continue;
                    }
                    if (dialog.RequestedAction == TepDialogAction.ExportSettings)
                    {
                        WriteJson(dialog.RequestedPath, dialog.SelectedSettings);
                        TaskDialog.Show("Расчёт ТЭП", "Настройки экспортированы.");
                        continue;
                    }
                    if (dialog.RequestedAction == TepDialogAction.Clear)
                    {
                        ClearPluginData(doc, LoadResult(doc));
                        TaskDialog.Show("Расчёт ТЭП", "Сохранённый результат и созданные командой проверочные объекты удалены.");
                        continue;
                    }
                    if (dialog.RequestedAction != TepDialogAction.Calculate) continue;

                    settings = dialog.SelectedSettings;
                    SaveSettings(doc, settings);
                    sources = ScanSources(doc, settings);
                    TepCalculationOutput output = Calculate(doc, sources, settings);

                    if (settings.CreateViews || settings.Create3D || settings.CreateSchedule)
                        CreateRevitOutput(doc, output, LoadResult(doc));

                    output.Status = output.Warnings.Any(x => x.Severity == "Ошибка по объекту") ? "Неполный расчёт" : output.Warnings.Count > 0 ? "Успешно с предупреждениями" : "Успешно";

                    SaveResult(doc, output);
                    ExportRequested(output);
                    if (settings.ShowResults) ShowResultDialog(uidoc, output);
                    else TaskDialog.Show("Расчёт ТЭП", string.Format("{0}\nСтрок результатов: {1}\nПредупреждений: {2}", output.Status, output.Results.Count, output.Warnings.Count));
                    return Result.Succeeded;
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Критическая ошибка расчёта ТЭП", ex.ToString());
                return Result.Failed;
            }
        }

        private static void ShowResultDialog(UIDocument uidoc, TepCalculationOutput output)
        {
            while (true)
            {
                AR_calculateTEP resultDialog = new AR_calculateTEP(uidoc, output);
                if (resultDialog.ShowDialog() != true) return;
                if (resultDialog.RequestedAction == TepDialogAction.ExportCsv) ExportCsv(output, resultDialog.RequestedPath);
                else if (resultDialog.RequestedAction == TepDialogAction.ExportXlsx) ExportXlsx(output, resultDialog.RequestedPath);
                else return;
            }
        }

        private static List<SourceContext> ScanSources(Document host, TepSettings settings)
        {
            List<SourceContext> result = new List<SourceContext>();
            TepSourceOption hostOption = new TepSourceOption { Key = "HOST", Name = host.Title, Kind = "Текущая модель", Status = "Доступна", IsIncluded = true, LinkInstanceId = -1 };
            ApplySourceSelection(hostOption, settings);
            result.Add(new SourceContext { Option = hostOption, Document = host, Transform = Transform.Identity });

            foreach (RevitLinkInstance link in new FilteredElementCollector(host).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>())
            {
                string key = "LINK:" + IDHelper.ElIdValue(link.Id).ToString(CultureInfo.InvariantCulture);
                Document linked = null;
                string status;
                try { linked = link.GetLinkDocument(); status = linked == null ? "Выгружена / недоступна" : "Доступна"; }
                catch (Exception ex) { status = "Ошибка: " + ex.Message; }
                TepSourceOption option = new TepSourceOption { Key = key, Name = link.Name, Kind = "Revit-связь", Status = status, IsIncluded = linked != null, LinkInstanceId = IDHelper.ElIdValue(link.Id) };
                ApplySourceSelection(option, settings);
                if (linked == null) option.IsIncluded = false;
                Transform transform = Transform.Identity;
                try { transform = link.GetTotalTransform(); } catch { }
                result.Add(new SourceContext { Option = option, Document = linked, Transform = transform });
            }
            return result;
        }

        private static void ApplySourceSelection(TepSourceOption option, TepSettings settings)
        {
            if (settings == null || settings.IncludedSourceKeys == null || settings.IncludedSourceKeys.Count == 0) return;
            option.IsIncluded = settings.IncludedSourceKeys.Contains(option.Key);
            option.GraphicsOnly = settings.GraphicsOnlySourceKeys != null && settings.GraphicsOnlySourceKeys.Contains(option.Key);
        }

        private static TepCalculationOutput Calculate(Document host, IList<SourceContext> sources, TepSettings settings)
        {
            TepCalculationOutput output = new TepCalculationOutput
            {
                CalculationId = DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture),
                Date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
                User = Environment.UserName,
                Method = (settings.Method == "Old" ? "Старая" : "Новая") + " / " + settings.RulesVersion,
                Settings = settings
            };
            List<SpatialRecord> basis = new List<SpatialRecord>();
            List<SpatialRecord> rooms = new List<SpatialRecord>();
            List<ParkingRecord> parking = new List<ParkingRecord>();

            foreach (SourceContext source in sources)
            {
                if (source.Document == null)
                {
                    AddWarning(output, "LINK-001", "Предупреждение", "Связь выгружена или недоступна.", "Все", source.Option.Name, string.Empty, "RevitLinkInstance", source.Option.LinkInstanceId, source.Option.LinkInstanceId, "Загрузить и проверить связь.", "Исключена из расчёта");
                    continue;
                }
                if (!source.Option.IsIncluded || source.Option.GraphicsOnly) continue;
                try
                {
                    List<SpatialRecord> sourceRooms = CollectRooms(source, settings, output);
                    rooms.AddRange(sourceRooms);
                    if (string.Equals(settings.AreaBasis, "Areas", StringComparison.OrdinalIgnoreCase)) basis.AddRange(CollectAreas(source, settings, output));
                    else basis.AddRange(sourceRooms);
                    parking.AddRange(CollectParking(source, settings, output));
                }
                catch (Exception ex)
                {
                    AddWarning(output, "SRC-001", "Ошибка по объекту", "Источник не обработан: " + ex.Message, "Все", source.Option.Name, string.Empty, "Документ", -1, source.Option.LinkInstanceId, "Проверить файл и права доступа.", "Исключён из расчёта");
                }
            }

            string[] selectedBuildings = SplitWords(settings.SelectedBuildings);
            if (selectedBuildings.Length > 0)
            {
                basis = basis.Where(x => selectedBuildings.Any(b => string.Equals(b, x.Building, StringComparison.OrdinalIgnoreCase))).ToList();
                rooms = rooms.Where(x => selectedBuildings.Any(b => string.Equals(b, x.Building, StringComparison.OrdinalIgnoreCase))).ToList();
                parking = parking.Where(x => selectedBuildings.Any(b => string.Equals(b, x.Building, StringComparison.OrdinalIgnoreCase))).ToList();
            }
            if (basis.Count == 0) AddWarning(output, "DATA-001", "Предупреждение", "Не найдено пространственных объектов выбранного типа с положительной площадью в выбранных границах.", "Площади", host.Title, string.Empty, settings.AreaBasis, -1, -1, "Проверить границы помещений/зон, фильтр корпусов и основу расчёта.", "Результат неполный");
            ValidateDuplicates(rooms, output);
            AddMethodWarning(output, settings, host.Title);
            BuildResults(output, basis, rooms, parking, settings);
            BuildDetails(output, basis, rooms, parking, settings);
            output.Status = output.Warnings.Any(x => x.Severity == "Ошибка по объекту") ? "Неполный расчёт" : output.Warnings.Count > 0 ? "Успешно с предупреждениями" : "Успешно";
            return output;
        }

        private static List<SpatialRecord> CollectRooms(SourceContext source, TepSettings settings, TepCalculationOutput output)
        {
            List<SpatialRecord> result = new List<SpatialRecord>();
            foreach (Room room in new FilteredElementCollector(source.Document).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().Cast<Room>())
            {
                double area = SafeArea(room);
                if (room.Location == null || area <= 1e-9)
                {
                    AddWarning(output, "ROOM-001", "Ошибка по объекту", "Помещение не размещено, не замкнуто или имеет нулевую площадь.", "Площади", source.Option.Name, Building(room, source, settings), "Помещение", IDHelper.ElIdValue(room.Id), source.Option.LinkInstanceId, "Разместить помещение и проверить границы.", "Исключено из расчёта");
                    continue;
                }
                result.Add(CreateSpatialRecord(room, source, settings, output, area, "Помещение"));
            }
            return result;
        }

        private static List<SpatialRecord> CollectAreas(SourceContext source, TepSettings settings, TepCalculationOutput output)
        {
            List<SpatialRecord> result = new List<SpatialRecord>();
            foreach (Area areaElement in new FilteredElementCollector(source.Document).OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType().Cast<Area>())
            {
                double area = SafeArea(areaElement);
                if (areaElement.Location == null || area <= 1e-9)
                {
                    AddWarning(output, "AREA-001", "Ошибка по объекту", "Зона не размещена, не замкнута или имеет нулевую площадь.", "Площади", source.Option.Name, Building(areaElement, source, settings), "Зона", IDHelper.ElIdValue(areaElement.Id), source.Option.LinkInstanceId, "Разместить зону и проверить границы.", "Исключена из расчёта");
                    continue;
                }
                result.Add(CreateSpatialRecord(areaElement, source, settings, output, area, "Зона"));
            }
            return result;
        }

        private static SpatialRecord CreateSpatialRecord(SpatialElement element, SourceContext source, TepSettings settings, TepCalculationOutput output, double area, string category)
        {
            string name = GetBuiltInString(element, BuiltInParameter.ROOM_NAME);
            if (string.IsNullOrWhiteSpace(name)) name = element.Name;
            string function = FirstNotEmpty(GetParameterString(element, settings.FunctionParameter), GetBuiltInString(element, BuiltInParameter.ROOM_DEPARTMENT), name);
            string levelName = string.Empty; double elevation = 0;
            Level level = source.Document.GetElement(element.LevelId) as Level;
            if (level != null) { levelName = level.Name; elevation = ToMeters(level.Elevation); }
            string building = Building(element, source, settings);
            if (string.IsNullOrWhiteSpace(building))
            {
                building = "Не определён";
                AddWarning(output, "BLDG-001", "Предупреждение", "Не определён корпус.", "Разрез по корпусам", source.Option.Name, building, category, IDHelper.ElIdValue(element.Id), source.Option.LinkInstanceId, "Заполнить параметр корпуса или изменить способ группировки.", "Учтено условно");
            }
            bool technical = ContainsAny(function, settings.TechnicalWords);
            bool residential = ContainsAny(function, settings.ResidentialWords);
            bool publicUse = ContainsAny(function, settings.PublicWords);
            if (!technical && !residential && !publicUse)
                AddWarning(output, "CLASS-001", "Предупреждение", "Не распознано функциональное назначение: " + function, "Классификация", source.Option.Name, building, category, IDHelper.ElIdValue(element.Id), source.Option.LinkInstanceId, "Настроить словарь или параметр назначения.", "Учтено условно");
            bool summer = IsTrue(element, settings.SummerParameter, settings.FalseWords) || ContainsAny(function, settings.SummerWords);
            double coefficient = summer ? ReadDouble(element, settings.SummerCoefficientParameter, DefaultSummerCoefficient(function)) : 1.0;
            double height = element is Room ? ToMeters(((Room)element).UnboundedHeight) : 3.0;
            double volume = 0; bool approximate = false;
            Parameter volumeParameter = element.get_Parameter(BuiltInParameter.ROOM_VOLUME);
            if (volumeParameter != null && volumeParameter.StorageType == StorageType.Double && volumeParameter.AsDouble() > 0) volume = ToCubicMeters(volumeParameter.AsDouble());
            if (volume <= 0) { volume = area * Math.Max(0.1, height); approximate = true; }
            string includeText = GetParameterString(element, settings.IncludeParameter);
            bool include = string.IsNullOrWhiteSpace(includeText) || !ContainsAny(includeText, settings.FalseWords);
            bool underground = elevation < settings.ZeroElevationMeters - 0.001 || ContainsAny(levelName + " " + function, settings.UndergroundWords);
            return new SpatialRecord
            {
                SourceKey = source.Option.Key,
                Source = source.Option.Name,
                HostLinkId = source.Option.LinkInstanceId,
                ElementId = IDHelper.ElIdValue(element.Id),
                Building = building,
                Level = levelName,
                Category = category,
                ElementName = FirstNotEmpty(GetBuiltInString(element, BuiltInParameter.ROOM_NUMBER), element.Name),
                Function = function,
                Apartment = GetParameterString(element, settings.ApartmentParameter),
                BuildingType = GetParameterString(element, settings.BuildingTypeParameter),
                Area = area,
                Volume = volume,
                Elevation = elevation,
                Height = height,
                SummerCoefficient = Math.Max(0, coefficient),
                Include = include,
                Residential = residential,
                Public = publicUse,
                Technical = technical,
                Summer = summer,
                Underground = underground,
                BuiltIn = IsTrue(element, settings.BuiltInParameter, settings.FalseWords),
                Standalone = IsTrue(element, settings.StandaloneParameter, settings.FalseWords),
                ApproximateVolume = approximate
            };
        }

        private static List<ParkingRecord> CollectParking(SourceContext source, TepSettings settings, TepCalculationOutput output)
        {
            List<ParkingRecord> result = new List<ParkingRecord>();
            foreach (FamilyInstance element in new FilteredElementCollector(source.Document).OfClass(typeof(FamilyInstance)).WhereElementIsNotElementType().Cast<FamilyInstance>())
            {
                bool categoryParking = element.Category != null && IDHelper.ElIdValue(element.Category.Id) == (long)BuiltInCategory.OST_Parking;
                if (!categoryParking && !IsTrue(element, settings.ParkingParameter, settings.FalseWords)) continue;
                Level level = source.Document.GetElement(element.LevelId) as Level;
                string levelName = level == null ? string.Empty : level.Name;
                double elevation = level == null ? 0 : ToMeters(level.Elevation);
                bool underground = elevation < settings.ZeroElevationMeters - 0.001 || ContainsAny(levelName, settings.UndergroundWords);
                if (!underground) AddWarning(output, "PARK-001", "Предупреждение", "Машино-место не отнесено к подземной части.", "15", source.Option.Name, Building(element, source, settings), "Машино-место", IDHelper.ElIdValue(element.Id), source.Option.LinkInstanceId, "Проверить уровень или признак подземной части.", "Исключено из показателя 15");
                result.Add(new ParkingRecord { SourceKey = source.Option.Key, Source = source.Option.Name, HostLinkId = source.Option.LinkInstanceId, ElementId = IDHelper.ElIdValue(element.Id), Building = Building(element, source, settings), Level = levelName, Name = element.Name, Underground = underground });
            }
            return result;
        }

        private static void ValidateDuplicates(List<SpatialRecord> rooms, TepCalculationOutput output)
        {
            foreach (IGrouping<string, SpatialRecord> group in rooms.Where(x => !string.IsNullOrWhiteSpace(x.Apartment)).GroupBy(x => x.SourceKey + "|" + x.Building + "|" + x.Apartment, StringComparer.OrdinalIgnoreCase))
            {
                if (group.Select(x => x.Source).Distinct().Count() > 1) AddWarning(output, "APT-002", "Предупреждение", "Идентификатор квартиры встречается в нескольких источниках: " + group.First().Apartment, "14", group.First().Source, group.First().Building, "Помещение", group.First().ElementId, group.First().HostLinkId, "Проверить идентификаторы квартир.", "Требуется ручная проверка");
            }
            foreach (SpatialRecord room in rooms.Where(x => (x.Residential || ContainsAny(x.Function, "квартир")) && string.IsNullOrWhiteSpace(x.Apartment)))
                AddWarning(output, "APT-001", "Предупреждение", "Квартирное помещение без ID/номера квартиры.", "11, 12, 14", room.Source, room.Building, room.Category, room.ElementId, room.HostLinkId, "Заполнить параметр квартиры.", "Исключено из квартирных показателей");
        }

        private static void AddMethodWarning(TepCalculationOutput output, TepSettings settings, string source)
        {
            AddWarning(output, "RULES-001", "Предупреждение", "Использован редактируемый профиль формул версии " + settings.RulesVersion + ". До утверждения нормативного приложения формулы следует считать проектными настройками, а не окончательно согласованной методикой.", "1–17", source, "Все", "Методика", -1, -1, "Проверить формулы на вкладке «Методика» и утвердить версию правил.", "Требуется ручная проверка");
        }

        private static void BuildResults(TepCalculationOutput output, List<SpatialRecord> basis, List<SpatialRecord> rooms, List<ParkingRecord> parking, TepSettings settings)
        {
            AddResultSlice(output, basis, rooms, parking, settings, "Все", "Все");
            foreach (string building in basis.Select(x => x.Building).Concat(parking.Select(x => x.Building)).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x))
                AddResultSlice(output, basis.Where(x => x.Building == building).ToList(), rooms.Where(x => x.Building == building).ToList(), parking.Where(x => x.Building == building).ToList(), settings, building, "Все");
            foreach (string source in basis.Select(x => x.Source).Concat(parking.Select(x => x.Source)).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x))
                AddResultSlice(output, basis.Where(x => x.Source == source).ToList(), rooms.Where(x => x.Source == source).ToList(), parking.Where(x => x.Source == source).ToList(), settings, "Все", source);
        }

        private static void AddResultSlice(TepCalculationOutput output, List<SpatialRecord> basis, List<SpatialRecord> rooms, List<ParkingRecord> parking, TepSettings settings, string building, string source)
        {
            for (int i = 0; i < Indicators.GetLength(0); ++i)
            {
                string number = Indicators[i, 0];
                if (settings.SelectedIndicators != null && !settings.SelectedIndicators.Contains(number)) continue;
                MetricValue value = Metric(number, basis, rooms, parking, settings);
                double correctedValue; string correctionReason;
                if (building == "Все" && source == "Все" && TryGetManualCorrection(settings.ManualCorrections, number, out correctedValue, out correctionReason))
                {
                    output.Corrections.Add(new TepCorrectionRow { Date = output.Date, Author = output.User, Indicator = number, Reason = correctionReason, OldValue = value.Value, NewValue = correctedValue });
                    output.Details.Add(new TepDetailRow { Indicator = number, Source = "Ручная корректировка", Building = "Все", Category = "Корректировка", Element = output.User, ElementId = -1, HostLinkId = -1, RawValue = correctedValue, Unit = Indicators[i, 2], Rule = "Заменено " + value.Value.ToString("R", CultureInfo.InvariantCulture) + " → " + correctedValue.ToString("R", CultureInfo.InvariantCulture) + "; причина: " + correctionReason });
                    value.Value = correctedValue; value.Status = "Ручная корректировка"; value.Comment = correctionReason;
                }
                output.Results.Add(new TepResultRow { Number = number, Name = Indicators[i, 1], RawValue = value.Value, DisplayValue = Math.Round(value.Value, settings.RoundingDigits, MidpointRounding.AwayFromZero).ToString("F" + settings.RoundingDigits, CultureInfo.CurrentCulture), Unit = Indicators[i, 2], Building = building, Source = source, Method = output.Method, Status = value.Status, Comment = value.Comment });
            }
        }

        private static MetricValue Metric(string number, List<SpatialRecord> basis, List<SpatialRecord> rooms, List<ParkingRecord> parking, TepSettings settings)
        {
            MetricValue core = MetricCore(number, basis, rooms, parking, settings);
            IList<TepFormulaRule> rules = settings.Method == "Old" ? settings.OldMethodRules : settings.NewMethodRules;
            TepFormulaRule rule = rules == null ? null : rules.FirstOrDefault(x => x != null && string.Equals(x.Number, number, StringComparison.OrdinalIgnoreCase));
            if (rule == null || string.IsNullOrWhiteSpace(rule.Expression)) return core;
            try
            {
                Dictionary<string, double> variables = FormulaVariables(basis, rooms, parking, settings);
                core.Value = new FormulaParser(rule.Expression, variables).Evaluate();
                string formulaComment = "Формула профиля: " + rule.Expression;
                core.Comment = string.IsNullOrWhiteSpace(core.Comment) ? formulaComment : core.Comment + " " + formulaComment;
                return core;
            }
            catch (Exception ex)
            {
                core.Status = "Ошибка формулы; использовано значение ядра";
                core.Comment = "Формула «" + rule.Expression + "»: " + ex.Message;
                return core;
            }
        }

        private static Dictionary<string, double> FormulaVariables(List<SpatialRecord> basis, List<SpatialRecord> rooms, List<ParkingRecord> parking, TepSettings settings)
        {
            return new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase)
            {
                {"A_ALL", MetricCore("1", basis, rooms, parking, settings).Value},
                {"A_RES_BUILDING", MetricCore("2", basis, rooms, parking, settings).Value},
                {"A_RES_IN_RES_BUILDING", MetricCore("2а", basis, rooms, parking, settings).Value},
                {"A_NONRES_IN_RES_BUILDING", MetricCore("2б", basis, rooms, parking, settings).Value},
                {"A_NONRES_BUILDING", MetricCore("3", basis, rooms, parking, settings).Value},
                {"FOOTPRINT", MetricCore("4", basis, rooms, parking, settings).Value},
                {"V_ALL", MetricCore("5", basis, rooms, parking, settings).Value},
                {"V_ABOVE", MetricCore("5а", basis, rooms, parking, settings).Value},
                {"V_BELOW", MetricCore("5б", basis, rooms, parking, settings).Value},
                {"A_ABOVE", MetricCore("6а", basis, rooms, parking, settings).Value},
                {"A_BELOW", MetricCore("6б", basis, rooms, parking, settings).Value},
                {"A_NNP_BUILTIN", MetricCore("7", basis, rooms, parking, settings).Value},
                {"A_NNP_STANDALONE", MetricCore("8", basis, rooms, parking, settings).Value},
                {"A_ABOVE_RES_BUILDING", MetricCore("9а", basis, rooms, parking, settings).Value},
                {"A_ABOVE_NONRES_BUILDING", MetricCore("9б", basis, rooms, parking, settings).Value},
                {"A_PUBLIC_CALC", MetricCore("10", basis, rooms, parking, settings).Value},
                {"A_APT_COEF", MetricCore("11", basis, rooms, parking, settings).Value},
                {"A_APT_HEATED", MetricCore("12", basis, rooms, parking, settings).Value},
                {"A_PUBLIC", MetricCore("13", basis, rooms, parking, settings).Value},
                {"COUNT_APT", MetricCore("14", basis, rooms, parking, settings).Value},
                {"COUNT_PARKING", MetricCore("15", basis, rooms, parking, settings).Value},
                {"FLOORS_ABOVE", MetricCore("16", basis, rooms, parking, settings).Value},
                {"FLOORS_ALL", MetricCore("17", basis, rooms, parking, settings).Value}
            };
        }

        private static MetricValue MetricCore(string number, List<SpatialRecord> basis, List<SpatialRecord> rooms, List<ParkingRecord> parking, TepSettings settings)
        {
            List<SpatialRecord> valid = basis.Where(x => x.Include && (settings.Method == "Old" || !x.Technical)).ToList();
            Func<SpatialRecord, bool> residentialBuilding = x => settings.BuildingType == "Residential" || (settings.BuildingType == "ByParameters" && (x.Residential || ContainsAny(x.BuildingType, settings.ResidentialWords)));
            Func<SpatialRecord, bool> nonResidentialBuilding = x => settings.BuildingType == "Public" || settings.BuildingType == "HighRise" || (settings.BuildingType == "ByParameters" && !residentialBuilding(x));
            if (number == "1") return V(valid.Sum(x => x.Area));
            if (number == "2") return V(valid.Where(residentialBuilding).Sum(x => x.Area));
            if (number == "2а") return V(valid.Where(x => residentialBuilding(x) && x.Residential).Sum(x => x.Area));
            if (number == "2б") return V(valid.Where(x => residentialBuilding(x) && !x.Residential).Sum(x => x.Area));
            if (number == "3") return V(valid.Where(nonResidentialBuilding).Sum(x => x.Area));
            if (number == "4") return Footprint(valid, settings);
            if (number == "5" || number == "5а" || number == "5б")
            {
                IEnumerable<SpatialRecord> volumeRecords = valid;
                if (number == "5а") volumeRecords = volumeRecords.Where(x => !x.Underground);
                if (number == "5б") volumeRecords = volumeRecords.Where(x => x.Underground);
                List<SpatialRecord> list = volumeRecords.ToList();
                return new MetricValue { Value = list.Sum(x => x.Volume), Status = list.Any(x => x.ApproximateVolume) ? "Требует проверки" : "Рассчитано", Comment = list.Any(x => x.ApproximateVolume) ? "Часть объёмов рассчитана как площадь × высота помещения." : string.Empty };
            }
            if (number == "6") return V(valid.Sum(x => x.Area));
            if (number == "6а" || number == "9") return V(valid.Where(x => !x.Underground).Sum(x => x.Area));
            if (number == "6б") return V(valid.Where(x => x.Underground).Sum(x => x.Area));
            if (number == "7") return V(valid.Where(x => x.BuiltIn && !x.Standalone && (x.Public || !x.Residential)).Sum(x => x.Area));
            if (number == "8") return V(valid.Where(x => x.Standalone && (x.Public || !x.Residential)).Sum(x => x.Area));
            if (number == "9а") return V(valid.Where(x => !x.Underground && residentialBuilding(x)).Sum(x => x.Area));
            if (number == "9б") return V(valid.Where(x => !x.Underground && nonResidentialBuilding(x)).Sum(x => x.Area));
            if (number == "10") return V(valid.Where(x => x.Public && !ContainsAny(x.Function, settings.ExcludedPublicWords)).Sum(x => x.Area));
            List<SpatialRecord> apartmentRooms = rooms.Where(x => x.Include && !string.IsNullOrWhiteSpace(x.Apartment)).ToList();
            if (number == "11") return V(apartmentRooms.Sum(x => x.Summer ? x.Area * x.SummerCoefficient : x.Area));
            if (number == "12") return V(apartmentRooms.Where(x => !x.Summer).Sum(x => x.Area));
            if (number == "13") return V(rooms.Where(x => x.Include && x.Public).Sum(x => x.Area));
            if (number == "14") return V(apartmentRooms.Select(x => x.SourceKey + "|" + x.Building + "|" + x.Apartment).Distinct(StringComparer.OrdinalIgnoreCase).Count());
            if (number == "15") return V(parking.Count(x => x.Underground));
            if (number == "16") return V(MaxFloorCount(valid.Where(x => !x.Underground), settings));
            if (number == "17") return V(MaxFloorCount(valid, settings));
            return V(0);
        }

        private sealed class FormulaParser
        {
            private readonly string _text;
            private readonly IDictionary<string, double> _variables;
            private int _position;

            public FormulaParser(string text, IDictionary<string, double> variables)
            {
                _text = (text ?? string.Empty).Trim();
                _variables = variables;
                if (_text.Length == 0) throw new InvalidOperationException("пустая формула");
                if (_text.Length > 500) throw new InvalidOperationException("формула длиннее 500 символов");
            }

            public double Evaluate()
            {
                double value = ParseExpression();
                SkipSpaces();
                if (_position != _text.Length) throw new InvalidOperationException("неожиданный символ в позиции " + (_position + 1).ToString(CultureInfo.InvariantCulture));
                if (double.IsNaN(value) || double.IsInfinity(value)) throw new InvalidOperationException("получено недопустимое число");
                return value;
            }

            private double ParseExpression()
            {
                double value = ParseTerm();
                while (true)
                {
                    SkipSpaces();
                    if (Take('+')) value += ParseTerm();
                    else if (Take('-')) value -= ParseTerm();
                    else return value;
                }
            }

            private double ParseTerm()
            {
                double value = ParseFactor();
                while (true)
                {
                    SkipSpaces();
                    if (Take('*')) value *= ParseFactor();
                    else if (Take('/'))
                    {
                        double divisor = ParseFactor();
                        if (Math.Abs(divisor) < 1e-12) throw new DivideByZeroException("деление на ноль");
                        value /= divisor;
                    }
                    else return value;
                }
            }

            private double ParseFactor()
            {
                SkipSpaces();
                if (Take('+')) return ParseFactor();
                if (Take('-')) return -ParseFactor();
                if (Take('('))
                {
                    double nested = ParseExpression(); SkipSpaces();
                    if (!Take(')')) throw new InvalidOperationException("не закрыта скобка");
                    return nested;
                }
                if (_position >= _text.Length) throw new InvalidOperationException("формула неожиданно закончилась");
                if (char.IsDigit(_text[_position]) || _text[_position] == '.' || _text[_position] == ',') return ParseNumber();
                if (char.IsLetter(_text[_position]) || _text[_position] == '_') return ParseVariable();
                throw new InvalidOperationException("недопустимый символ «" + _text[_position] + "»");
            }

            private double ParseNumber()
            {
                int start = _position;
                while (_position < _text.Length && (char.IsDigit(_text[_position]) || _text[_position] == '.' || _text[_position] == ',')) _position++;
                double value;
                if (!double.TryParse(_text.Substring(start, _position - start).Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out value)) throw new InvalidOperationException("неверное число");
                return value;
            }

            private double ParseVariable()
            {
                int start = _position;
                while (_position < _text.Length && (char.IsLetterOrDigit(_text[_position]) || _text[_position] == '_')) _position++;
                string name = _text.Substring(start, _position - start); double value;
                if (!_variables.TryGetValue(name, out value)) throw new InvalidOperationException("неизвестная переменная «" + name + "»");
                return value;
            }

            private bool Take(char value) { if (_position < _text.Length && _text[_position] == value) { _position++; return true; } return false; }
            private void SkipSpaces() { while (_position < _text.Length && char.IsWhiteSpace(_text[_position])) _position++; }
        }

        private static MetricValue Footprint(List<SpatialRecord> records, TepSettings settings)
        {
            double total = 0;
            foreach (IGrouping<string, SpatialRecord> building in records.GroupBy(x => x.Building ?? string.Empty))
            {
                IGrouping<string, SpatialRecord> closest = building.GroupBy(x => x.Level ?? string.Empty).OrderBy(g => Math.Abs(g.Average(x => x.Elevation) - settings.GroundElevationMeters)).FirstOrDefault();
                if (closest != null) total += closest.Sum(x => x.Area);
            }
            return new MetricValue { Value = total, Status = "Требует проверки", Comment = "Оценка по сумме площадей пространств уровня, ближайшего к отметке земли; внешний контур следует проверить графически." };
        }

        private static double MaxFloorCount(IEnumerable<SpatialRecord> records, TepSettings settings)
        {
            string[] excluded = SplitWords(settings.ExcludedLevelWords);
            return records.Where(x => !ContainsAny(x.Level, excluded)).GroupBy(x => x.Building ?? string.Empty).Select(g => g.Select(x => x.Level).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).Count()).DefaultIfEmpty(0).Max();
        }

        private static MetricValue V(double value) { return new MetricValue { Value = value }; }

        private static void BuildDetails(TepCalculationOutput output, List<SpatialRecord> basis, List<SpatialRecord> rooms, List<ParkingRecord> parking, TepSettings settings)
        {
            foreach (SpatialRecord record in basis)
            {
                output.Details.Add(new TepDetailRow { Indicator = "База площадей/объёмов", Source = record.Source, Building = record.Building, Level = record.Level, Category = record.Category, Element = record.ElementName, ElementId = record.ElementId, HostLinkId = record.HostLinkId, Function = record.Function, Apartment = record.Apartment, RawValue = record.Area, Unit = "м²", Rule = record.Include ? (record.Technical ? "Техническое; исключается из основных сумм" : "Включено") : "Исключено параметром" });
                output.Details.Add(new TepDetailRow { Indicator = "5/5а/5б", Source = record.Source, Building = record.Building, Level = record.Level, Category = record.Category, Element = record.ElementName, ElementId = record.ElementId, HostLinkId = record.HostLinkId, Function = record.Function, Apartment = record.Apartment, RawValue = record.Volume, Unit = "м³", Rule = record.ApproximateVolume ? "Площадь × высота" : "Параметр объёма" });
            }
            foreach (SpatialRecord record in rooms.Where(x => !string.IsNullOrWhiteSpace(x.Apartment)))
                output.Details.Add(new TepDetailRow { Indicator = "11/12/14", Source = record.Source, Building = record.Building, Level = record.Level, Category = record.Category, Element = record.ElementName, ElementId = record.ElementId, HostLinkId = record.HostLinkId, Function = record.Function, Apartment = record.Apartment, RawValue = record.Summer ? record.Area * record.SummerCoefficient : record.Area, Unit = "м²", Rule = record.Summer ? "Летнее помещение; коэффициент " + record.SummerCoefficient.ToString(CultureInfo.InvariantCulture) : "Отапливаемое помещение" });
            foreach (ParkingRecord record in parking)
                output.Details.Add(new TepDetailRow { Indicator = "15", Source = record.Source, Building = record.Building, Level = record.Level, Category = "Машино-место", Element = record.Name, ElementId = record.ElementId, HostLinkId = record.HostLinkId, RawValue = record.Underground ? 1 : 0, Unit = "шт.", Rule = record.Underground ? "Подземное; включено" : "Наземное/неопределённое; исключено" });
        }

        private static void CreateRevitOutput(Document doc, TepCalculationOutput output, TepCalculationOutput previous)
        {
            using (Transaction transaction = new Transaction(doc, "KPLN. Расчёт ТЭП. Проверочные данные"))
            {
                transaction.Start();
                try
                {
                    if (output.Settings.UpdateGraphics && previous != null) DeleteCreated(doc, previous.CreatedElementIds);
                    if (output.Settings.CreateSchedule) CreateScheduleOrReport(doc, output);
                    if (output.Settings.CreateViews) { CreateDraftingReport(doc, output); CreatePlanViews(doc, output); }
                    if (output.Settings.Create3D) Create3DView(doc, output);
                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    transaction.RollBack();
                    AddWarning(output, "GRAPH-001", "Ошибка по объекту", "Проверочная графика не создана: " + ex.Message, "Графика", doc.Title, "Все", "Виды/спецификации", -1, -1, "Проверить права на запись и шаблоны видов.", "Расчётные результаты сохранены");
                }
            }
        }

        private static void CreateScheduleOrReport(Document doc, TepCalculationOutput output)
        {
            string name = output.Settings.Prefix + "Спецификация_" + output.CalculationId;
            ViewSchedule schedule = null;
            try
            {
                Category roomsCategory = Category.GetCategory(doc, BuiltInCategory.OST_Rooms);
                schedule = ViewSchedule.CreateKeySchedule(doc, roomsCategory.Id);
                schedule.Name = name;
                IList<SchedulableField> available = schedule.Definition.GetSchedulableFields();
                foreach (BuiltInParameter bip in new[] { BuiltInParameter.ROOM_DEPARTMENT, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS })
                {
                    SchedulableField field = available.FirstOrDefault(x => IDHelper.ElIdValue(x.ParameterId) == (long)bip);
                    if (field != null) schedule.Definition.AddField(field);
                }
                TableSectionData body = schedule.GetTableData().GetSectionData(SectionType.Body);
                foreach (TepResultRow row in output.Results.Where(x => x.Building == "Все" && x.Source == "Все"))
                {
                    int newRow = body.LastRowNumber + 1;
                    body.InsertRow(newRow);
                    if (body.NumberOfColumns > 0) body.SetCellText(newRow, 0, row.Number + ". " + row.Name);
                    if (body.NumberOfColumns > 1) body.SetCellText(newRow, 1, row.DisplayValue + " " + row.Unit);
                    if (body.NumberOfColumns > 2) body.SetCellText(newRow, 2, row.Status + (string.IsNullOrWhiteSpace(row.Comment) ? string.Empty : ": " + row.Comment));
                }
                output.CreatedElementIds.Add(IDHelper.ElIdValue(schedule.Id));
            }
            catch (Exception ex)
            {
                if (schedule != null && doc.GetElement(schedule.Id) != null) try { doc.Delete(schedule.Id); } catch { }
                AddWarning(output, "SCHED-001", "Предупреждение", "Key-спецификация Revit недоступна: " + ex.Message + ". Создан табличный отчёт на чертёжном виде.", "Отчёт", doc.Title, "Все", "Спецификация", -1, -1, "Проверить поддержку key schedule для помещений.", "Учтено условно");
                CreateDraftingReport(doc, output);
            }
        }

        private static void CreateDraftingReport(Document doc, TepCalculationOutput output)
        {
            ViewFamilyType type = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault(x => x.ViewFamily == ViewFamily.Drafting);
            if (type == null) throw new InvalidOperationException("Не найден тип чертёжного вида.");
            ViewDrafting view = ViewDrafting.Create(doc, type.Id);
            view.Name = output.Settings.Prefix + "Отчёт_" + output.CalculationId + "_" + output.CreatedElementIds.Count.ToString(CultureInfo.InvariantCulture);
            ElementId textTypeId = new FilteredElementCollector(doc).OfClass(typeof(TextNoteType)).FirstElementId();
            TextNoteOptions options = new TextNoteOptions(textTypeId);
            StringBuilder text = new StringBuilder();
            text.AppendLine("РАСЧЁТ ТЭП  " + output.CalculationId);
            text.AppendLine(output.Method + "  |  " + output.Date + "  |  " + output.Status);
            text.AppendLine(new string('─', 92));
            foreach (TepResultRow row in output.Results.Where(x => x.Building == "Все" && x.Source == "Все")) text.AppendLine(string.Format("{0,-3} {1,-58} {2,12} {3}", row.Number, Trim(row.Name, 58), row.DisplayValue, row.Unit));
            text.AppendLine(new string('─', 92));
            text.AppendLine("Предупреждений: " + output.Warnings.Count.ToString(CultureInfo.InvariantCulture));
            TextNote note = TextNote.Create(doc, view.Id, new XYZ(0, 0, 0), text.ToString(), options);
            output.CreatedElementIds.Add(IDHelper.ElIdValue(view.Id)); output.CreatedElementIds.Add(IDHelper.ElIdValue(note.Id));
        }

        private static void Create3DView(Document doc, TepCalculationOutput output)
        {
            ViewFamilyType type = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault(x => x.ViewFamily == ViewFamily.ThreeDimensional);
            if (type == null) throw new InvalidOperationException("Не найден тип 3D-вида.");
            View3D view = View3D.CreateIsometric(doc, type.Id);
            view.Name = output.Settings.Prefix + "3D_Проверка_" + output.CalculationId;
            ApplyOverrides(doc, view, output, null);
            output.CreatedElementIds.Add(IDHelper.ElIdValue(view.Id));
        }

        private static void CreatePlanViews(Document doc, TepCalculationOutput output)
        {
            Dictionary<string, ViewPlan> plans = new FilteredElementCollector(doc).OfClass(typeof(ViewPlan)).Cast<ViewPlan>()
                .Where(x => !x.IsTemplate && x.GenLevel != null && x.ViewType == ViewType.FloorPlan)
                .GroupBy(x => x.GenLevel.Name, StringComparer.OrdinalIgnoreCase).ToDictionary(x => x.Key, x => x.First(), StringComparer.OrdinalIgnoreCase);
            foreach (string level in output.Details.Where(x => x.HostLinkId <= 0 && !string.IsNullOrWhiteSpace(x.Level)).Select(x => x.Level).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                ViewPlan original;
                if (!plans.TryGetValue(level, out original)) continue;
                ElementId duplicateId;
                try { duplicateId = original.Duplicate(ViewDuplicateOption.Duplicate); } catch { continue; }
                ViewPlan duplicate = doc.GetElement(duplicateId) as ViewPlan;
                if (duplicate == null) continue;
                duplicate.Name = output.Settings.Prefix + "Проверка_" + level + "_" + output.CalculationId;
                ApplyOverrides(doc, duplicate, output, level);
                output.CreatedElementIds.Add(IDHelper.ElIdValue(duplicate.Id));
            }
        }

        private static void ApplyOverrides(Document doc, View view, TepCalculationOutput output, string level)
        {
            ElementId solidFill = new FilteredElementCollector(doc).OfClass(typeof(FillPatternElement)).Cast<FillPatternElement>().Where(x => x.GetFillPattern().IsSolidFill).Select(x => x.Id).FirstOrDefault();
            foreach (TepDetailRow row in output.Details.Where(x => x.HostLinkId <= 0 && x.ElementId > 0 && (level == null || string.Equals(x.Level, level, StringComparison.OrdinalIgnoreCase))).GroupBy(x => x.ElementId).Select(x => x.First()))
            {
                ElementId id = IDHelper.CreateElementId(row.ElementId); if (doc.GetElement(id) == null) continue;
                bool excluded = (row.Rule ?? string.Empty).IndexOf("исключ", StringComparison.OrdinalIgnoreCase) >= 0;
                Color color = excluded ? new Color(220, 70, 70) : new Color(70, 180, 90);
                OverrideGraphicSettings graphics = new OverrideGraphicSettings().SetProjectionLineColor(color).SetProjectionLineWeight(5);
                if (solidFill != null && solidFill != ElementId.InvalidElementId) graphics.SetSurfaceForegroundPatternId(solidFill).SetSurfaceForegroundPatternColor(color).SetSurfaceTransparency(55);
                try { view.SetElementOverrides(id, graphics); } catch { }
            }
            foreach (long linkId in output.Details.Where(x => x.HostLinkId > 0).Select(x => x.HostLinkId).Concat((output.Settings.GraphicsOnlySourceKeys ?? new List<string>()).Where(x => x.StartsWith("LINK:", StringComparison.OrdinalIgnoreCase)).Select(x => ParseLong(x.Substring(5)))).Where(x => x > 0).Distinct())
            {
                ElementId id = IDHelper.CreateElementId(linkId); if (doc.GetElement(id) == null) continue;
                try { view.SetElementOverrides(id, new OverrideGraphicSettings().SetProjectionLineColor(new Color(235, 150, 35)).SetProjectionLineWeight(4)); } catch { }
            }
        }

        private static long ParseLong(string value) { long result; return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out result) ? result : -1; }

        private static void DeleteCreated(Document doc, IEnumerable<long> ids)
        {
            if (ids == null) return;
            foreach (long value in ids.Distinct().Where(x => x > 0))
            {
                ElementId id = IDHelper.CreateElementId(value);
                if (doc.GetElement(id) != null) try { doc.Delete(id); } catch { }
            }
        }

        private static void ClearPluginData(Document doc, TepCalculationOutput previous)
        {
            using (Transaction transaction = new Transaction(doc, "KPLN. Очистка данных ТЭП"))
            {
                transaction.Start();
                DeleteCreated(doc, previous == null ? null : previous.CreatedElementIds);
                foreach (DataStorage storage in FindStorages(doc)) doc.Delete(storage.Id);
                transaction.Commit();
            }
        }

        private static void ExportRequested(TepCalculationOutput output)
        {
            if (output.Settings.ExportCsv)
            {
                SaveFileDialog d = new SaveFileDialog { Filter = "CSV (*.csv)|*.csv", FileName = "ТЭП_" + output.CalculationId + ".csv" };
                if (d.ShowDialog() == true) ExportCsv(output, d.FileName);
            }
            if (output.Settings.ExportXlsx)
            {
                SaveFileDialog d = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", FileName = "ТЭП_" + output.CalculationId + ".xlsx" };
                if (d.ShowDialog() == true) ExportXlsx(output, d.FileName);
            }
        }

        private static void ExportCsv(TepCalculationOutput output, string path)
        {
            StringBuilder sb = new StringBuilder();
            CsvSection(sb, "Итоги", new[] { "№", "Показатель", "Значение без округления", "Значение", "Ед.", "Корпус", "Источник", "Методика", "Статус", "Комментарий" }, output.Results.Select(x => new[] { x.Number, x.Name, x.RawValue.ToString("R", CultureInfo.InvariantCulture), x.DisplayValue, x.Unit, x.Building, x.Source, x.Method, x.Status, x.Comment }));
            CsvSection(sb, "Детализация", new[] { "Показатель", "Источник", "Корпус", "Уровень", "Категория", "Элемент", "ElementId", "Функция", "Квартира", "Значение", "Ед.", "Правило" }, output.Details.Select(x => new[] { x.Indicator, x.Source, x.Building, x.Level, x.Category, x.Element, x.ElementId.ToString(CultureInfo.InvariantCulture), x.Function, x.Apartment, x.RawValue.ToString("R", CultureInfo.InvariantCulture), x.Unit, x.Rule }));
            CsvSection(sb, "Предупреждения", new[] { "Код", "Критичность", "Описание", "Показатель", "Источник", "Корпус", "Категория", "ElementId", "Действие", "Учёт" }, output.Warnings.Select(x => new[] { x.Code, x.Severity, x.Description, x.Indicator, x.Source, x.Building, x.Category, x.ElementId.ToString(CultureInfo.InvariantCulture), x.UserAction, x.Accounting }));
            CsvSection(sb, "Настройки и методика", new[] { "Параметр", "Значение" }, SettingsRows(output));
            CsvSection(sb, "Журнал ручных корректировок", new[] { "Дата", "Автор", "Показатель", "Причина", "Старое", "Новое" }, (output.Corrections ?? new List<TepCorrectionRow>()).Select(x => new[] { x.Date, x.Author, x.Indicator, x.Reason, x.OldValue.ToString("R", CultureInfo.InvariantCulture), x.NewValue.ToString("R", CultureInfo.InvariantCulture) }));
            File.WriteAllText(path, sb.ToString(), new UTF8Encoding(true));
        }

        private static void CsvSection(StringBuilder sb, string title, string[] headers, IEnumerable<string[]> rows)
        {
            sb.AppendLine("[" + title + "]"); sb.AppendLine(string.Join(";", headers.Select(Csv)));
            foreach (string[] row in rows) sb.AppendLine(string.Join(";", row.Select(Csv)));
            sb.AppendLine();
        }

        private static string Csv(string value) { return "\"" + (value ?? string.Empty).Replace("\"", "\"\"") + "\""; }

        private static void ExportXlsx(TepCalculationOutput output, string path)
        {
            List<Tuple<string, List<string[]>>> sheets = new List<Tuple<string, List<string[]>>>();
            Func<IEnumerable<TepResultRow>, List<string[]>> resultRows = rows => Header(new[] { "№", "Показатель", "Значение", "Ед.", "Корпус", "Источник", "Методика", "Статус", "Комментарий" }, rows.Select(x => new[] { x.Number, x.Name, x.RawValue.ToString("R", CultureInfo.InvariantCulture), x.Unit, x.Building, x.Source, x.Method, x.Status, x.Comment }));
            sheets.Add(Tuple.Create("Итоги", resultRows(output.Results.Where(x => x.Building == "Все" && x.Source == "Все"))));
            sheets.Add(Tuple.Create("По корпусам", resultRows(output.Results.Where(x => x.Building != "Все" && x.Source == "Все"))));
            sheets.Add(Tuple.Create("По связям", resultRows(output.Results.Where(x => x.Source != "Все" && x.Building == "Все"))));
            sheets.Add(Tuple.Create("По уровням", Header(new[] { "Источник", "Корпус", "Уровень", "Площадь, м²" }, output.Details.Where(x => x.Unit == "м²").GroupBy(x => new { x.Source, x.Building, x.Level }).Select(g => new[] { g.Key.Source, g.Key.Building, g.Key.Level, g.Sum(x => x.RawValue).ToString("R", CultureInfo.InvariantCulture) }))));
            sheets.Add(Tuple.Create("Детализация площадей", DetailRows(output.Details.Where(x => x.Unit == "м²"))));
            sheets.Add(Tuple.Create("Детализация объёмов", DetailRows(output.Details.Where(x => x.Unit == "м³"))));
            sheets.Add(Tuple.Create("Квартиры", DetailRows(output.Details.Where(x => x.Indicator.Contains("11")))));
            sheets.Add(Tuple.Create("Машино-места", DetailRows(output.Details.Where(x => x.Indicator == "15"))));
            sheets.Add(Tuple.Create("Предупреждения", Header(new[] { "Код", "Критичность", "Описание", "Показатель", "Источник", "Корпус", "Категория", "ElementId", "Действие", "Учёт" }, output.Warnings.Select(x => new[] { x.Code, x.Severity, x.Description, x.Indicator, x.Source, x.Building, x.Category, x.ElementId.ToString(CultureInfo.InvariantCulture), x.UserAction, x.Accounting }))));
            sheets.Add(Tuple.Create("Настройки и методика", Header(new[] { "Параметр", "Значение" }, SettingsRows(output))));
            sheets.Add(Tuple.Create("Журнал корректировок", Header(new[] { "Дата", "Автор", "Показатель", "Причина", "Старое", "Новое" }, (output.Corrections ?? new List<TepCorrectionRow>()).Select(x => new[] { x.Date, x.Author, x.Indicator, x.Reason, x.OldValue.ToString("R", CultureInfo.InvariantCulture), x.NewValue.ToString("R", CultureInfo.InvariantCulture) }))));
            WriteWorkbook(path, sheets);
        }

        private static List<string[]> DetailRows(IEnumerable<TepDetailRow> details)
        {
            return Header(new[] { "Показатель", "Источник", "Корпус", "Уровень", "Категория", "Элемент", "ElementId", "Функция", "Квартира", "Значение", "Ед.", "Правило" }, details.Select(x => new[] { x.Indicator, x.Source, x.Building, x.Level, x.Category, x.Element, x.ElementId.ToString(CultureInfo.InvariantCulture), x.Function, x.Apartment, x.RawValue.ToString("R", CultureInfo.InvariantCulture), x.Unit, x.Rule }));
        }

        private static IEnumerable<string[]> SettingsRows(TepCalculationOutput output)
        {
            yield return new[] { "ID расчёта", output.CalculationId }; yield return new[] { "Дата", output.Date }; yield return new[] { "Пользователь", output.User }; yield return new[] { "Методика", output.Method };
            yield return new[] { "Тип объекта", output.Settings.BuildingType }; yield return new[] { "Основа площадей", output.Settings.AreaBasis }; yield return new[] { "Группировка", output.Settings.Grouping };
            yield return new[] { "0.000, м", output.Settings.ZeroElevationMeters.ToString(CultureInfo.InvariantCulture) }; yield return new[] { "Земля, м", output.Settings.GroundElevationMeters.ToString(CultureInfo.InvariantCulture) };
            yield return new[] { "Параметр корпуса", output.Settings.BuildingParameter }; yield return new[] { "Параметр назначения", output.Settings.FunctionParameter }; yield return new[] { "Параметр квартиры", output.Settings.ApartmentParameter };
            yield return new[] { "Округление", output.Settings.RoundingDigits.ToString(CultureInfo.InvariantCulture) };
            IList<TepFormulaRule> rules = output.Settings.Method == "Old" ? output.Settings.OldMethodRules : output.Settings.NewMethodRules;
            if (rules != null) foreach (TepFormulaRule rule in rules.Where(x => x != null)) yield return new[] { "Формула " + rule.Number + " — " + rule.Name, rule.Expression };
        }

        private static List<string[]> Header(string[] header, IEnumerable<string[]> rows) { List<string[]> result = new List<string[]> { header }; result.AddRange(rows); return result; }

        private static void WriteWorkbook(string path, IList<Tuple<string, List<string[]>>> sheets)
        {
            if (File.Exists(path)) File.Delete(path);
            using (Package package = Package.Open(path, FileMode.Create, FileAccess.ReadWrite))
            {
                Uri workbookUri = PackUriHelper.CreatePartUri(new Uri("/xl/workbook.xml", UriKind.Relative));
                PackagePart workbook = package.CreatePart(workbookUri, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", CompressionOption.Maximum);
                package.CreateRelationship(new Uri("xl/workbook.xml", UriKind.Relative), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
                StringBuilder workbookXml = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>");
                for (int i = 0; i < sheets.Count; ++i)
                {
                    string relId = "rId" + (i + 1).ToString(CultureInfo.InvariantCulture);
                    Uri sheetUri = PackUriHelper.CreatePartUri(new Uri("/xl/worksheets/sheet" + (i + 1).ToString(CultureInfo.InvariantCulture) + ".xml", UriKind.Relative));
                    PackagePart sheet = package.CreatePart(sheetUri, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", CompressionOption.Maximum);
                    workbook.CreateRelationship(new Uri("worksheets/sheet" + (i + 1).ToString(CultureInfo.InvariantCulture) + ".xml", UriKind.Relative), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", relId);
                    WritePart(sheet, SheetXml(sheets[i].Item2));
                    workbookXml.Append("<sheet name=\"").Append(XmlEscape(SafeSheetName(sheets[i].Item1))).Append("\" sheetId=\"").Append(i + 1).Append("\" r:id=\"").Append(relId).Append("\"/>");
                }
                workbookXml.Append("</sheets></workbook>"); WritePart(workbook, workbookXml.ToString());
            }
        }

        private static string SheetXml(IList<string[]> rows)
        {
            StringBuilder sb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");
            for (int r = 0; r < rows.Count; ++r)
            {
                sb.Append("<row r=\"").Append(r + 1).Append("\">");
                for (int c = 0; c < rows[r].Length; ++c) sb.Append("<c r=\"").Append(ColumnName(c + 1)).Append(r + 1).Append("\" t=\"inlineStr\"><is><t xml:space=\"preserve\">").Append(XmlEscape(rows[r][c])).Append("</t></is></c>");
                sb.Append("</row>");
            }
            return sb.Append("</sheetData></worksheet>").ToString();
        }

        private static void WritePart(PackagePart part, string content) { using (Stream stream = part.GetStream(FileMode.Create, FileAccess.Write)) using (StreamWriter writer = new StreamWriter(stream, new UTF8Encoding(false))) writer.Write(content); }
        private static string ColumnName(int n) { string s = string.Empty; while (n > 0) { n--; s = (char)('A' + n % 26) + s; n /= 26; } return s; }
        private static string SafeSheetName(string name) { string s = new string((name ?? "Лист").Where(c => "[]:*?/\\".IndexOf(c) < 0).ToArray()); return s.Length > 31 ? s.Substring(0, 31) : s; }
        private static string XmlEscape(string value) { return System.Security.SecurityElement.Escape(value ?? string.Empty) ?? string.Empty; }

        private static string Building(Element element, SourceContext source, TepSettings settings)
        {
            if (settings.Grouping == "Link") return source.Option.Key == "HOST" ? source.Document.Title : source.Option.Name;
            if (settings.Grouping == "Workset")
            {
                try { Workset workset = source.Document.GetWorksetTable().GetWorkset(element.WorksetId); return workset == null ? string.Empty : workset.Name; } catch { return string.Empty; }
            }
            return GetParameterString(element, settings.BuildingParameter);
        }

        private static double SafeArea(SpatialElement element)
        {
            Parameter p = element.get_Parameter(BuiltInParameter.ROOM_AREA);
            return p == null || p.StorageType != StorageType.Double ? 0 : ToSquareMeters(p.AsDouble());
        }

        private static string GetBuiltInString(Element element, BuiltInParameter parameter)
        {
            Parameter p = element.get_Parameter(parameter); return ParameterText(p);
        }

        private static string GetParameterString(Element element, string name)
        {
            if (element == null || string.IsNullOrWhiteSpace(name)) return string.Empty;
            Parameter p = element.LookupParameter(name); if (p == null && element.Document != null)
            {
                Element type = element.Document.GetElement(element.GetTypeId()); if (type != null) p = type.LookupParameter(name);
            }
            return ParameterText(p);
        }

        private static string ParameterText(Parameter p)
        {
            if (p == null || !p.HasValue) return string.Empty;
            if (p.StorageType == StorageType.String) return p.AsString() ?? string.Empty;
            if (p.StorageType == StorageType.Integer) return p.AsInteger().ToString(CultureInfo.InvariantCulture);
            if (p.StorageType == StorageType.Double) return p.AsDouble().ToString("R", CultureInfo.InvariantCulture);
            if (p.StorageType == StorageType.ElementId) return IDHelper.ElIdValue(p.AsElementId()).ToString(CultureInfo.InvariantCulture);
            return p.AsValueString() ?? string.Empty;
        }

        private static bool IsTrue(Element element, string parameter, string falseWords)
        {
            string value = GetParameterString(element, parameter); if (string.IsNullOrWhiteSpace(value)) return false;
            string normalized = value.Trim();
            return !SplitWords(falseWords).Any(x => string.Equals(x, normalized, StringComparison.OrdinalIgnoreCase));
        }

        private static double ReadDouble(Element element, string parameter, double fallback)
        {
            Parameter p = string.IsNullOrWhiteSpace(parameter) ? null : element.LookupParameter(parameter);
            if (p != null && p.HasValue)
            {
                if (p.StorageType == StorageType.Double) return p.AsDouble();
                double parsed; if (double.TryParse(ParameterText(p).Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out parsed)) return parsed;
            }
            return fallback;
        }

        private static double DefaultSummerCoefficient(string function)
        {
            string s = (function ?? string.Empty).ToLowerInvariant(); if (s.Contains("балкон") || s.Contains("террас")) return 0.3; if (s.Contains("лоджи")) return 0.5; if (s.Contains("веранд")) return 1.0; return 0.3;
        }

        private static bool ContainsAny(string value, string csv) { return ContainsAny(value, SplitWords(csv)); }
        private static bool ContainsAny(string value, IEnumerable<string> words)
        {
            string source = (value ?? string.Empty).Trim(); return words.Any(x => !string.IsNullOrWhiteSpace(x) && source.IndexOf(x.Trim(), StringComparison.OrdinalIgnoreCase) >= 0);
        }
        private static string[] SplitWords(string csv) { return (csv ?? string.Empty).Split(new[] { ',', ';', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).Where(x => x.Length > 0).ToArray(); }
        private static bool TryGetManualCorrection(string text, string indicator, out double value, out string reason)
        {
            value = 0; reason = string.Empty;
            foreach (string rawLine in (text ?? string.Empty).Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string[] reasonParts = rawLine.Split(new[] { '|' }, 2); string[] valueParts = reasonParts[0].Split(new[] { '=' }, 2);
                if (valueParts.Length != 2 || !string.Equals(valueParts[0].Trim(), indicator, StringComparison.OrdinalIgnoreCase)) continue;
                if (!double.TryParse(valueParts[1].Trim().Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out value)) return false;
                reason = reasonParts.Length > 1 && !string.IsNullOrWhiteSpace(reasonParts[1]) ? reasonParts[1].Trim() : "Причина не указана";
                return true;
            }
            return false;
        }
        private static string FirstNotEmpty(params string[] values) { return values.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? string.Empty; }
        private static string Trim(string value, int length) { value = value ?? string.Empty; return value.Length <= length ? value : value.Substring(0, Math.Max(0, length - 1)) + "…"; }

        private static double ToMeters(double feet) { return UnitUtils.ConvertFromInternalUnits(feet, UnitTypeId.Meters); }
        private static double ToSquareMeters(double squareFeet) { return UnitUtils.ConvertFromInternalUnits(squareFeet, UnitTypeId.SquareMeters); }
        private static double ToCubicMeters(double cubicFeet) { return UnitUtils.ConvertFromInternalUnits(cubicFeet, UnitTypeId.CubicMeters); }

        private static void AddWarning(TepCalculationOutput output, string code, string severity, string description, string indicator, string source, string building, string category, long elementId, long hostLinkId, string action, string accounting)
        {
            output.Warnings.Add(new TepWarningRow { Code = code, Severity = severity, Description = description, Indicator = indicator, Source = source, Building = building, Category = category, ElementId = elementId, HostLinkId = hostLinkId, UserAction = action, Accounting = accounting });
        }

        private static Schema GetSchema()
        {
            Schema schema = Schema.Lookup(StorageSchemaGuid); if (schema != null) return schema;
            SchemaBuilder builder = new SchemaBuilder(StorageSchemaGuid); builder.SetSchemaName("KPLN_TEP_Data"); builder.SetReadAccessLevel(AccessLevel.Public); builder.SetWriteAccessLevel(AccessLevel.Public); builder.AddSimpleField(SettingsField, typeof(string)); builder.AddSimpleField(ResultField, typeof(string)); return builder.Finish();
        }

        private static IEnumerable<DataStorage> FindStorages(Document doc)
        {
            Schema schema = GetSchema(); return new FilteredElementCollector(doc).OfClass(typeof(DataStorage)).Cast<DataStorage>().Where(x => x.Name == StorageName && x.GetEntity(schema).IsValid()).ToList();
        }

        private static DataStorage GetOrCreateStorage(Document doc)
        {
            DataStorage storage = FindStorages(doc).FirstOrDefault(); if (storage != null) return storage;
            storage = DataStorage.Create(doc); storage.Name = StorageName; return storage;
        }

        private static TepSettings LoadSettings(Document doc) { string json = LoadField(doc, SettingsField); return string.IsNullOrWhiteSpace(json) ? null : Deserialize<TepSettings>(json); }
        private static TepCalculationOutput LoadResult(Document doc) { string json = LoadField(doc, ResultField); return string.IsNullOrWhiteSpace(json) ? null : Deserialize<TepCalculationOutput>(json); }
        private static string LoadField(Document doc, string field)
        {
            DataStorage storage = FindStorages(doc).FirstOrDefault(); if (storage == null) return null; Entity entity = storage.GetEntity(GetSchema()); return entity.IsValid() ? entity.Get<string>(field) : null;
        }

        private static void SaveSettings(Document doc, TepSettings settings) { SaveField(doc, SettingsField, Serialize(settings), "KPLN. Настройки ТЭП"); }
        private static void SaveResult(Document doc, TepCalculationOutput output) { SaveField(doc, ResultField, Serialize(output), "KPLN. Результат ТЭП"); }
        private static void SaveField(Document doc, string field, string value, string transactionName)
        {
            using (Transaction transaction = new Transaction(doc, transactionName))
            {
                transaction.Start(); Schema schema = GetSchema(); DataStorage storage = GetOrCreateStorage(doc); Entity entity = storage.GetEntity(schema); if (!entity.IsValid()) entity = new Entity(schema); entity.Set<string>(field, value ?? string.Empty); storage.SetEntity(entity); transaction.Commit();
            }
        }

        private static string Serialize<T>(T value)
        {
            DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(T)); using (MemoryStream ms = new MemoryStream()) { serializer.WriteObject(ms, value); return Encoding.UTF8.GetString(ms.ToArray()); }
        }
        private static T Deserialize<T>(string json)
        {
            DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(T)); using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(json))) return (T)serializer.ReadObject(ms);
        }
        private static void WriteJson<T>(string path, T value) { File.WriteAllText(path, Serialize(value), new UTF8Encoding(true)); }
        private static T ReadJson<T>(string path) { return Deserialize<T>(File.ReadAllText(path, Encoding.UTF8)); }
    }
}