using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using KPLN_Tools.Common;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Command_AR_CalculateTEP : IExternalCommand
    {
        public Result Execute(ExternalCommandData data, ref string message, ElementSet elements)
        {
            if (data.Application.ActiveUIDocument == null) return Result.Cancelled;
            try
            {
                var engine = new Engine(data.Application.ActiveUIDocument);
                var window = new KPLN_Tools.Forms.AR_CalculateTEP(engine);
                new System.Windows.Interop.WindowInteropHelper(window).Owner = data.Application.MainWindowHandle;
                window.ShowDialog(); return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return Result.Cancelled; }
            catch (Exception ex) { message = ex.ToString(); return Result.Failed; }
        }
        public static readonly string[] Titles = {
            "1. Суммарная поэтажная площадь всего", "2.1. Суммарная поэтажная площадь жилых зданий",
            "3.1. Жилая часть СПП жилых зданий", "3.2. Нежилая часть СПП жилых зданий",
            "2.2. Суммарная поэтажная площадь нежилых зданий",
            "4. Общая площадь квартир с летними помещениями и коэффициентами",
            "5. Площадь квартир без летних помещений, балконов и террас",
            "6. Площадь помещений общественного назначения", "7. Количество квартир" };
        public static readonly string[] Roles = { "Включить", "Исключить", "Летнее помещение", "Шахта", "Проём", "Многосветное пространство", "Лестничный просвет", "Терраса / эксплуатируемая кровля", "Французский балкон" };
        public const string MethodDescription =
            "Проёмы: старая методика учитывает их на одном этаже только при площади более 36 м²; новая - при любой площади.\n\n" +
            "Шахты: в старой методике нет автоматического ограничения одним этажом; в новой они учитываются только на нижнем учитываемом этаже.\n\n" +
            "Террасы и эксплуатируемая кровля: в старой методике учитываются по настройкам включения; в новой исключаются из ГНС.";
        public class Rule
        {
            public string Role { get; set; }
            public string Category { get; set; }
            public string Parameter { get; set; }
            public string Operation { get; set; }
            public string Value { get; set; }
            public double Factor { get; set; }
            public string Label
            {
                get
                {
                    return Role + ": " + (string.IsNullOrEmpty(Category) ? "все категории" : Category) +
                (string.IsNullOrEmpty(Parameter) ? "" : " · " + Parameter + " " + Operation + " «" + Value + "»") +
                (Role == "Летнее помещение" ? " × " + Factor.ToString("0.###") : "");
                }
            }
            public Rule Copy() { return (Rule)MemberwiseClone(); }
        }
        public class Metric
        {
            public int Index { get; set; }
            public string Title { get { return Titles[Index]; } }
            public bool Enabled { get; set; }
            public string Primary { get; set; }
            public string LevelKey { get; set; }
            public bool AboveOnly { get; set; }
            public bool AllIncludes { get; set; }
            public string AreaScheme { get; set; }
            public string Phase { get; set; }
            public string ApartmentParameter { get; set; }
            public string SectionParameter { get; set; }
            public string VoidParameter { get; set; }
            public string FactorParameter { get; set; }
            public bool CountInstances { get; set; }
            public ObservableCollection<Rule> Rules { get; set; }
            public Metric() { Enabled = true; Primary = "Помещения"; AboveOnly = true; Rules = new ObservableCollection<Rule>(); }
            public Metric Copy(int index)
            { var c = (Metric)MemberwiseClone(); c.Index = index; c.Rules = new ObservableCollection<Rule>(Rules.Select(r => r.Copy())); return c; }
        }
        public class Settings
        {
            public string Method { get; set; }
            public string Grouping { get; set; }
            public string BuildingParameter { get; set; }
            public string ManualBuilding { get; set; }
            public bool SelectedOnly { get; set; }
            public string Ground { get; set; }
            public ObservableCollection<Metric> Old { get; set; }
            public ObservableCollection<Metric> New { get; set; }
            public Settings() { Method = "Новая"; Grouping = "По связям Revit"; ManualBuilding = "Корпус 1"; Ground = "0"; Old = Defaults(); New = Defaults(); }
            private static ObservableCollection<Metric> Defaults()
            { return new ObservableCollection<Metric>(Enumerable.Range(0, 9).Select(i => new Metric { Index = i, AboveOnly = i < 5 })); }
            public ObservableCollection<Metric> Metrics { get { return Method == "Старая" ? Old : New; } }
        }
        public class Source
        {
            public string Key { get; set; }
            public string Name { get; set; }
            public string Mode { get; set; }
            public Document Document { get; set; }
            public Transform Transform { get; set; }
            public ElementId RootLink { get; set; }
            public List<Element> Elements { get; set; }
            public bool Loaded { get { return Document != null; } }
        }
        public class LevelChoice
        {
            public string Key { get; set; }
            public string Name { get; set; }
            public double Elevation { get; set; }
        }
        public class Item
        {
            public Source Source;
            public Element Element;
            public string Building, LevelName, Apartment, Role, Reason;
            public double Elevation, Area;
            public Solid Shape;
            public double Factor = 1;
            public bool Excluded;
            public Item AtLevel(string name, double elevation)
            { var copy = (Item)MemberwiseClone(); copy.LevelName = name; copy.Elevation = elevation; return copy; }
            public string Identity { get { return Source.Key + "/" + Element.UniqueId; } }
            public string Description { get { return Source.Name + " | " + Building + " | " + LevelName + " | ID " + IDHelper.ElIdValue(Element.Id) + " | " + Element.Name; } }
        }
        public class Piece
        {
            public Solid Shape;
            public double Elevation;
            public string Building, Source, LevelName, Description;
            public double Factor = 1;
            public bool Excluded;
        }
        public class ResultRow
        {
            public Metric Metric;
            public double Value;
            public bool Incomplete, NeedsReview;
            public List<Item> Items = new List<Item>();
            public List<Piece> Pieces = new List<Piece>();
            public List<string> Messages = new List<string>();
            public Dictionary<string, double> ApartmentTotals = new Dictionary<string, double>();
            public void Error(string text) { Incomplete = true; Messages.Add("ОШИБКА: " + text); }
            public void Review(string text) { NeedsReview = true; Messages.Add("ПРОВЕРИТЬ: " + text); }
        }
        public class FloorTotal
        {
            public string Building { get; set; }
            public string LevelName { get; set; }
            public double Elevation { get; set; }
            public double Area { get; set; }
        }
        public static List<FloorTotal> SummarizeFloors(IEnumerable<FloorTotal> contributions)
        {
            // Same physical floor may have different names in different links.
            // Keep buildings separate and sort by elevation, not by the text of the level name.
            return contributions.GroupBy(p => new { p.Building, Elevation = Math.Round(p.Elevation, 6) })
                .Select(g => new FloorTotal
                {
                    Building = g.Key.Building,
                    Elevation = g.Key.Elevation,
                    LevelName = string.Join(" / ", g.Select(p => p.LevelName).Distinct().OrderBy(n => n)),
                    Area = g.Sum(p => p.Area)
                }).OrderBy(p => p.Building).ThenBy(p => p.Elevation).ToList();
        }
        public class Run
        {
            public string Method, Configuration;
            public DateTime Date = DateTime.Now;
            public List<ResultRow> Results = new List<ResultRow>();
            public List<string> Messages = new List<string>();
            public string GraphicsStatus = "Проверочная графика пока не создана.";
            public string Report(bool detailed = false)
            {
                var b = new StringBuilder();
                b.AppendLine("РАСЧЁТ ТЭП · " + Date.ToString("g") + " · " + Method + " методика · правила 1.0");
                if (detailed)
                {
                    b.AppendLine(Configuration); b.AppendLine(GraphicsStatus);
                    foreach (string s in Messages) b.AppendLine(s);
                }
                else
                {
                    b.AppendLine("Площади после объединения контуров и вычитания исключений. Этажи отсортированы по отметке.");
                    if (GraphicsStatus.IndexOf("ОШИБКА", StringComparison.OrdinalIgnoreCase) >= 0)
                        b.AppendLine("При построении проверочных контуров возникли ошибки. См. «Подробности и ошибки».");
                }
                if (Results.Count == 0) b.AppendLine("Не выбран ни один показатель.");
                foreach (var r in Results)
                {
                    b.AppendLine(); b.AppendLine(r.Metric.Title);
                    b.AppendLine("ИТОГО: " + r.Value.ToString(r.Metric.Index == 8 ? "0" : "N2") + (r.Metric.Index == 8 ? " шт." : " м²") +
                        " - " + (r.Incomplete ? "НЕПОЛНЫЙ РЕЗУЛЬТАТ" : r.NeedsReview ? "ТРЕБУЕТ ПРОВЕРКИ" : "Рассчитано по заданным правилам"));
                    if (r.Metric.Index != 8)
                    {
                        var floors = SummarizeFloors(r.Pieces.Where(p => !p.Excluded).Select(p => new FloorTotal
                        {
                            Building = p.Building,
                            LevelName = p.LevelName,
                            Elevation = p.Elevation,
                            Area = AreaM2(p.Shape) * p.Factor
                        }));
                        b.AppendLine();
                        b.AppendLine("ПЛОЩАДЬ КАЖДОГО ЭТАЖА:");
                        if (floors.Count == 0) b.AppendLine("  Нет рассчитанных площадей этажей. См. подробности и ошибки.");
                        foreach (var building in floors.GroupBy(p => p.Building))
                        {
                            b.AppendLine();
                            b.AppendLine("  Корпус: " + building.Key);
                            foreach (var floor in building)
                                b.AppendLine("    " + floor.LevelName + " · отм. " + (floor.Elevation * .3048).ToString("+0.000;-0.000;0.000") + " м - " + floor.Area.ToString("N2") + " м²");
                            b.AppendLine("  Итого по корпусу: " + building.Sum(p => p.Area).ToString("N2") + " м²");
                        }
                        b.AppendLine();
                        b.AppendLine("  Сумма по этажам: " + floors.Sum(p => p.Area).ToString("N2") + " м²");
                        b.AppendLine("  Показаны этажи с рассчитанным вкладом. Округление - только при выводе.");
                    }
                    if (!detailed)
                    {
                        int count = r.Messages.Distinct().Count();
                        if (count > 0) b.AppendLine("  Замечаний: " + count + ". Причины - во вкладке «Подробности и ошибки».");
                        continue;
                    }
                    b.AppendLine("  Источник: " + r.Metric.Primary + "; стадия: " + r.Metric.Phase + "; схема зон: " + r.Metric.AreaScheme);
                    foreach (var rule in r.Metric.Rules) b.AppendLine("  Правило: " + rule.Label);
                    foreach (var s in r.Messages.Distinct()) b.AppendLine("  " + s);
                    if (r.Metric.Index != 8)
                        foreach (var g in r.Pieces.Where(p => !p.Excluded).GroupBy(p => p.Building + " | " + p.Source + " | " + p.LevelName))
                            b.AppendLine("  " + g.Key + ": " + g.Sum(p => AreaM2(p.Shape) * p.Factor).ToString("N2") + " м²");
                    foreach (var apt in r.ApartmentTotals.OrderBy(x => x.Key))
                        b.AppendLine("  Квартира " + apt.Key + ": " + apt.Value.ToString(r.Metric.Index == 8 ? "0" : "N2") + (r.Metric.Index == 8 ? " шт." : " м²"));
                    b.AppendLine("  Исходные контуры (до объединения/вычитания):");
                    foreach (var i in r.Items)
                        b.AppendLine("  " + (i.Excluded ? "- " : "+ ") + i.Description + " | " + i.Area.ToString("N3") + " м² × " + i.Factor.ToString("0.###") +
                            (i.Excluded ? " | " + i.Reason : "") + (string.IsNullOrEmpty(i.Apartment) ? "" : " | квартира " + i.Apartment));
                }
                return b.ToString();
            }
        }
        public class Engine
        {
            private sealed class FailureLog : IFailuresPreprocessor
            {
                private readonly List<string> messages;
                public FailureLog(List<string> target) { messages = target; }
                public FailureProcessingResult PreprocessFailures(FailuresAccessor accessor)
                {
                    bool hasError = false;
                    foreach (var failure in accessor.GetFailureMessages())
                    {
                        messages.Add("Revit: " + failure.GetDescriptionText());
                        if (failure.GetSeverity() == FailureSeverity.Warning) accessor.DeleteWarning(failure);
                        else hasError = true;
                    }
                    return hasError ? FailureProcessingResult.ProceedWithRollBack : FailureProcessingResult.Continue;
                }
            }
            private readonly Document doc;
            private readonly HashSet<long> selection;
            private readonly List<string> scanMessages = new List<string>();
            public ObservableCollection<Source> Sources { get; private set; }
            public List<LevelChoice> Levels { get; private set; }
            public List<string> Categories { get; private set; }
            public List<string> Parameters { get; private set; }
            public List<string> Phases { get; private set; }
            public List<string> AreaSchemes { get; private set; }
            public Settings Settings { get; private set; }
            public Run LastRun { get; private set; }
            private const string AppId = "KPLN.CalculateTEP.v1";
            private static readonly Guid SchemaId = new Guid("3ba929d8-9344-44b1-a90a-c8e94a9046a4");
            public Engine(UIDocument ui)
            {
                doc = ui.Document;
                selection = new HashSet<long>(ui.Selection.GetElementIds().Select(IDHelper.ElIdValue));
                Sources = new ObservableCollection<Source>(); Levels = new List<LevelChoice>();
                AddSource(doc, Transform.Identity, "host", "Текущая модель: " + doc.Title, ElementId.InvalidElementId, new HashSet<string>());
                var categories = new SortedSet<string>(StringComparer.CurrentCultureIgnoreCase);
                var parameters = new SortedSet<string>(StringComparer.CurrentCultureIgnoreCase);
                var phases = new SortedSet<string>(); var schemes = new SortedSet<string>();
                foreach (var s in Sources.Where(x => x.Loaded))
                {
                    foreach (Phase p in s.Document.Phases) phases.Add(p.Name);
                    foreach (AreaScheme a in new FilteredElementCollector(s.Document).OfClass(typeof(AreaScheme))) schemes.Add(a.Name);
                    foreach (Level level in new FilteredElementCollector(s.Document).OfClass(typeof(Level)))
                        Levels.Add(new LevelChoice
                        {
                            Key = s.Key + "/" + level.UniqueId,
                            Name = s.Name + " · " + level.Name + " (" + (s.Transform.OfPoint(new XYZ(0, 0, level.ProjectElevation)).Z * .3048).ToString("0.000") + " м)",
                            Elevation = s.Transform.OfPoint(new XYZ(0, 0, level.ProjectElevation)).Z
                        });
                    var types = new HashSet<long>();
                    foreach (var e in s.Elements)
                    {
                        if (e.Category != null) categories.Add(e.Category.Name);
                        foreach (Parameter p in e.Parameters) if (p.Definition != null) parameters.Add(p.Definition.Name);
                        if (types.Add(IDHelper.ElIdValue(e.GetTypeId())))
                        {
                            var t = s.Document.GetElement(e.GetTypeId());
                            if (t != null) foreach (Parameter p in t.Parameters) if (p.Definition != null) parameters.Add(p.Definition.Name);
                        }
                    }
                }
                Categories = categories.ToList(); Parameters = parameters.ToList(); Phases = phases.ToList(); AreaSchemes = schemes.ToList();
                Levels = Levels.OrderBy(x => x.Elevation).ToList(); Settings = LoadSettings();
            }
            private void AddSource(Document model, Transform transform, string key, string name, ElementId root, HashSet<string> ancestors)
            {
                var s = new Source { Document = model, Transform = transform, Key = key, Name = name, Mode = "Включать", RootLink = root };
                Sources.Add(s);
                if (model == null) { scanMessages.Add("Связь недоступна: " + name); return; }
                string identity = model.PathName + "|" + model.Title;
                if (ancestors.Contains(identity)) { s.Mode = "Исключить"; s.Elements = new List<Element>(); scanMessages.Add("Циклическая связь: " + name); return; }
                s.Elements = new FilteredElementCollector(model).WhereElementIsNotElementType().ToElements()
                    .Where(e => !e.ViewSpecific && e.Category != null && (e.Category.CategoryType == CategoryType.Model || e is SpatialElement)
                        && !(e is RevitLinkInstance) && !(e is DirectShape && ((DirectShape)e).ApplicationId == AppId)).ToList();
                var next = new HashSet<string>(ancestors) { identity };
                foreach (RevitLinkInstance link in new FilteredElementCollector(model).OfClass(typeof(RevitLinkInstance)))
                {
                    try
                    {
                        var type = model.GetElement(link.GetTypeId()) as RevitLinkType;
                        if (key != "host" && type != null && type.AttachmentType == AttachmentType.Overlay) continue;
                        AddSource(link.GetLinkDocument(), transform.Multiply(link.GetTotalTransform()), key + "/" + link.UniqueId,
                            name + " → " + link.Name + " [" + IDHelper.ElIdValue(link.Id) + "]", key == "host" ? link.Id : root, next);
                    }
                    catch (Exception ex) { scanMessages.Add("Не удалось прочитать связь " + link.Name + ": " + ex.Message); }
                }
            }
            public List<string> Values(string category, string parameter)
            {
                var values = new SortedSet<string>(StringComparer.CurrentCultureIgnoreCase);
                foreach (var s in Sources.Where(x => x.Loaded && x.Mode != "Исключить"))
                    foreach (var e in s.Elements.Where(e => string.IsNullOrEmpty(category) || e.Category.Name == category))
                    { string problem; var value = Read(e, parameter, out problem); if (!string.IsNullOrWhiteSpace(value)) values.Add(value); }
                return values.ToList();
            }
            private bool InScope(Source s, Element e)
            { return !Settings.SelectedOnly || (s.Key == "host" ? selection.Contains(IDHelper.ElIdValue(e.Id)) : selection.Contains(IDHelper.ElIdValue(s.RootLink))); }
            private static bool IsPrimary(Element e, Metric m)
            {
                if (m.Primary == "Помещения") return e is Room;
                if (m.Primary == "Пространства") return e is Space;
                if (m.Primary == "Зоны") return e is Area && (string.IsNullOrWhiteSpace(m.AreaScheme) || ((Area)e).AreaScheme.Name == m.AreaScheme);
                return !(e is SpatialElement);
            }
            private static bool Relevant(Element e, Metric m)
            {
                if (e is Area && !string.IsNullOrWhiteSpace(m.AreaScheme) && ((Area)e).AreaScheme.Name != m.AreaScheme) return false;
                return IsPrimary(e, m) || m.Rules.Any(r => !string.IsNullOrEmpty(r.Category) && e.Category.Name == r.Category);
            }
            private static bool PhaseMatches(Element e, Metric m, out string issue)
            {
                issue = null;
                if (e.DesignOption != null && !e.DesignOption.IsPrimary) return false;
                Phase phase = e.Document.Phases.Cast<Phase>().FirstOrDefault(p => p.Name == m.Phase);
                if (phase == null) { issue = "В источнике отсутствует стадия «" + m.Phase + "»."; return false; }
                if (e is Room || e is Space)
                { var p = e.get_Parameter(BuiltInParameter.ROOM_PHASE); return p != null && p.AsElementId() == phase.Id; }
                if (e is Area || !e.HasPhases()) return true;
                var status = e.GetPhaseStatus(phase.Id);
                return status == ElementOnPhaseStatus.New || status == ElementOnPhaseStatus.Existing;
            }
            public string CheckParameters(Metric m)
            {
                var b = new StringBuilder("Проверка: " + m.Title + "\n");
                foreach (var r in m.Rules)
                {
                    int total = 0, found = 0, matched = 0; var problems = new List<string>();
                    foreach (var s in Sources.Where(x => x.Loaded && x.Mode != "Исключить"))
                        foreach (var e in s.Elements.Where(e => InScope(s, e) && Relevant(e, m) && (string.IsNullOrEmpty(r.Category) || e.Category.Name == r.Category)))
                        {
                            total++; string issue; Read(e, r.Parameter, out issue);
                            if (string.IsNullOrEmpty(r.Parameter) || issue == null) found++;
                            else problems.Add(s.Name + " / ID " + IDHelper.ElIdValue(e.Id) + ": " + issue);
                            string matchIssue; if (Matches(e, r, out matchIssue)) matched++;
                        }
                    b.AppendLine(r.Label + "\n  Элементов: " + total + "; параметр доступен: " + found + "; совпало: " + matched);
                    foreach (string p in problems) b.AppendLine("  " + p);
                }
                foreach (var name in new[] { Settings.BuildingParameter, m.ApartmentParameter, m.SectionParameter, m.VoidParameter, m.FactorParameter }.Where(x => !string.IsNullOrWhiteSpace(x)).Distinct())
                {
                    int total = 0, filled = 0;
                    foreach (var s in Sources.Where(x => x.Loaded && x.Mode != "Исключить"))
                        foreach (var e in s.Elements.Where(e => InScope(s, e) && Relevant(e, m)))
                        { total++; string issue; if (!string.IsNullOrWhiteSpace(Read(e, name, out issue)) && issue == null) filled++; }
                    b.AppendLine("Параметр «" + name + "»: заполнен однозначно у " + filled + " из " + total + " элементов.");
                }
                if (m.Rules.Count == 0) b.AppendLine("Правила не заданы. Только показатель 1 допускает весь основной источник.");
                return b.ToString();
            }
            public Run Calculate()
            {
                var run = new Run { Method = Settings.Method };
                run.Configuration = "Группировка: " + Settings.Grouping + "; параметр корпуса: " + Settings.BuildingParameter +
                    "; отметка земли (справочно): " + Settings.Ground + " м. Ключ квартиры: экземпляр связи + корпус + секция + ID. Основной вариант проекта; альтернативные варианты исключены.";
                run.Messages.AddRange(scanMessages);
                foreach (var s in Sources) run.Messages.Add(s.Name + ": " + s.Mode + (s.Loaded ? "" : " - НЕДОСТУПНА"));
                foreach (var m in Settings.Metrics.Where(x => x.Enabled))
                {
                    var r = new ResultRow { Metric = m.Copy(m.Index) }; run.Results.Add(r);
                    try { CalculateMetric(r); }
                    catch (Exception ex) { r.Error("Расчёт показателя прерван: " + ex.Message); }
                    if (Sources.Any(s => !s.Loaded && s.Mode == "Включать" &&
                        (!Settings.SelectedOnly || selection.Contains(IDHelper.ElIdValue(s.RootLink))))) r.Error("Часть включённых связей недоступна.");
                }
                Reconcile(run, 0, 1, 4); Reconcile(run, 1, 2, 3);
                if (run.Results.Count == 0) run.Messages.Add("Не выбран ни один показатель.");
                LastRun = run; return run;
            }
            private void CalculateMetric(ResultRow r)
            {
                var m = r.Metric; bool gns = m.Index < 5;
                if (Settings.SelectedOnly && selection.Count == 0) { r.Error("До запуска команды не выбрано ни одного элемента/экземпляра связи."); return; }
                if (m.Index != 0 && !m.Rules.Any(x => x.Role == "Включить")) { r.Error("Задайте хотя бы одно правило «Включить» для этого показателя."); return; }
                if (m.Primary == "Элементы" && !m.Rules.Any(x => x.Role == "Включить" && !string.IsNullOrEmpty(x.Category)))
                { r.Error("Для основного источника «Элементы» задайте включаемую категорию."); return; }
                if (string.IsNullOrWhiteSpace(m.Phase)) { r.Error("Выберите стадию."); return; }
                if (m.Primary == "Зоны" && string.IsNullOrWhiteSpace(m.AreaScheme)) { r.Error("Выберите одну схему зон."); return; }
                var first = Levels.FirstOrDefault(l => l.Key == m.LevelKey);
                if (m.AboveOnly && first == null) { r.Error("Выберите первый учитываемый наземный уровень."); return; }
                if (first != null && m.AboveOnly) r.Messages.Add("Нижний уровень: " + first.Name);
                if ((m.Index == 5 || m.Index == 6 || m.Index == 8 && !m.CountInstances) && string.IsNullOrWhiteSpace(m.ApartmentParameter))
                { r.Error("Укажите параметр ID квартиры."); return; }
                if (m.Index == 8 && m.CountInstances && m.Primary != "Элементы")
                { r.Error("Для подсчёта экземпляров квартир выберите источник «Элементы» и категорию семейства квартир."); return; }
                foreach (var s in Sources.Where(x => x.Loaded && x.Mode != "Исключить"))
                    foreach (var e in s.Elements.Where(e => InScope(s, e) && Relevant(e, m)))
                    {
                        try
                        {
                            string issue;
                            if (!PhaseMatches(e, m, out issue)) { if (issue != null) r.Error(s.Name + ": " + issue); continue; }
                            var includeRules = m.Rules.Where(x => x.Role == "Включить").ToList(); var matched = new List<Rule>();
                            foreach (var rule in m.Rules)
                            {
                                string problem;
                                if (Matches(e, rule, out problem)) matched.Add(rule);
                                else if (problem != null) r.Error(s.Name + " / ID " + IDHelper.ElIdValue(e.Id) + ": " + problem);
                            }
                            bool selected = includeRules.Count == 0 ? IsPrimary(e, m) : m.AllIncludes ? includeRules.All(matched.Contains) : includeRules.Any(matched.Contains);
                            bool excluded = matched.Any(x => x.Role == "Исключить");
                            var special = matched.Where(x => x.Role != "Включить" && x.Role != "Исключить").ToList();
                            if (!selected && !excluded && special.Count == 0) continue;
                            if (e is Area && string.IsNullOrWhiteSpace(m.AreaScheme)) { r.Error("Для включаемых зон укажите одну схему зон."); continue; }
                            if (special.Select(x => x.Role).Distinct().Count() > 1) { r.Error(s.Name + " / ID " + IDHelper.ElIdValue(e.Id) + ": конфликт типов помещения/пространства."); continue; }
                            var item = new Item
                            {
                                Source = s,
                                Element = e,
                                Role = special.Count > 0 ? special[0].Role : "Обычное",
                                Excluded = excluded || s.Mode == "Только проверка",
                                Reason = excluded ? "Правило исключения" : "Источник только для проверки"
                            };
                            var level = e.Document.GetElement(e.LevelId) as Level;
                            if (level == null)
                            {
                                var lp = e.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM) ?? e.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM)
                                    ?? e.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM) ?? e.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT);
                                if (lp != null && lp.StorageType == StorageType.ElementId) level = e.Document.GetElement(lp.AsElementId()) as Level;
                            }
                            if (level == null) { r.Error(item.Description + ": не определён уровень; элемент пропущен."); continue; }
                            item.LevelName = level.Name; item.Elevation = s.Transform.OfPoint(new XYZ(0, 0, level.ProjectElevation)).Z;
                            if (Math.Abs(s.Transform.BasisZ.DotProduct(XYZ.BasisZ) - 1) > 1e-7) { r.Error(item.Description + ": наклонённая связь не поддерживается."); continue; }
                            bool spanningShaft = e is Opening && IDHelper.ElIdValue(e.Category.Id) == (long)BuiltInCategory.OST_ShaftOpening;
                            if (!spanningShaft && m.AboveOnly && item.Elevation < first.Elevation - .003) continue;
                            item.Building = Building(item, r); if (string.IsNullOrWhiteSpace(item.Building)) continue;
                            if (!item.Excluded && (m.Index == 5 || m.Index == 6 || m.Index == 8))
                            {
                                if (m.Index == 8 && m.CountInstances) item.Apartment = item.Identity;
                                else
                                {
                                    item.Apartment = Read(e, m.ApartmentParameter, out issue);
                                    if (issue != null || string.IsNullOrWhiteSpace(item.Apartment)) { r.Error(item.Description + ": нет однозначного ID квартиры; элемент пропущен."); continue; }
                                }
                                string section = "";
                                if (!string.IsNullOrWhiteSpace(m.SectionParameter))
                                {
                                    section = Read(e, m.SectionParameter, out issue);
                                    if (issue != null || string.IsNullOrWhiteSpace(section)) { r.Error(item.Description + ": не заполнена секция квартиры."); continue; }
                                }
                                item.Apartment = Key(s.Key, item.Building, section, item.Apartment);
                            }
                            if (!item.Excluded && item.Role == "Летнее помещение")
                            {
                                if (special.Select(x => x.Factor).Distinct().Count() > 1) { r.Error(item.Description + ": несколько коэффициентов летнего помещения."); continue; }
                                item.Factor = special[0].Factor;
                                if (!string.IsNullOrWhiteSpace(m.FactorParameter))
                                {
                                    double factor;
                                    if (!ReadFactor(e, m.FactorParameter, out factor) || factor < 0 || factor > 1) { r.Error(item.Description + ": коэффициент должен быть безразмерным числом от 0 до 1."); continue; }
                                    item.Factor = factor;
                                }
                                if (item.Factor < 0 || item.Factor > 1 || double.IsNaN(item.Factor)) { r.Error(item.Description + ": недопустимый коэффициент."); continue; }
                            }
                            if (gns || m.Index == 7 || m.Index == 8) item.Factor = 1;
                            if (gns && Settings.Method == "Новая" && item.Role == "Терраса / эксплуатируемая кровля") { item.Excluded = true; item.Reason = "Новая методика: исключение из ГНС"; }
                            if ((m.Index == 5 || m.Index == 6) && item.Role == "Французский балкон") { item.Excluded = true; item.Reason = "Французский балкон"; }
                            if (m.Index == 6 && (item.Role == "Летнее помещение" || item.Role == "Терраса / эксплуатируемая кровля")) { item.Excluded = true; item.Reason = "Без летних помещений"; }
                            if (!gns && !selected && !excluded) continue;
                            // Special classification never includes an object contrary to the inclusion filters.
                            if (!selected && !item.Excluded) continue;
                            if (m.Index == 8)
                            {
                                if (!item.Excluded && m.CountInstances && !(e is FamilyInstance)) { r.Error(item.Description + ": ожидался экземпляр семейства квартиры."); continue; }
                                if (!m.CountInstances && e is SpatialElement && SpatialArea(e) <= 1e-9) { r.Error(item.Description + ": помещение не размещено или не замкнуто."); continue; }
                                try { item.Shape = Footprint(item, false, r); item.Area = AreaM2(item.Shape); }
                                catch (Exception ex) { r.Review(item.Description + ": нет проверочной геометрии: " + ex.Message); }
                            }
                            else { item.Shape = Footprint(item, gns, r); item.Area = AreaM2(item.Shape); }
                            if (spanningShaft && gns)
                            {
                                var extent = e.get_BoundingBox(null);
                                if (extent == null) throw new InvalidOperationException("Не определён диапазон этажей шахты; задайте поэтажные зоны.");
                                // Bounding box is used only for vertical extent, never to calculate area.
                                double bottom = extent.Min.Z, top = extent.Max.Z;
                                var levels = new FilteredElementCollector(s.Document).OfClass(typeof(Level)).Cast<Level>()
                                    .Where(l => l.ProjectElevation >= bottom - .003 && l.ProjectElevation <= top + .003).ToList();
                                foreach (var l in levels)
                                {
                                    double elevation = s.Transform.OfPoint(new XYZ(0, 0, l.ProjectElevation)).Z;
                                    if (!m.AboveOnly || elevation >= first.Elevation - .003) r.Items.Add(item.AtLevel(l.Name, elevation));
                                }
                                if (levels.Count == 0) r.Error(item.Description + ": шахта не пересекает ни одного уровня.");
                                r.Review(item.Description + ": диапазон уровней шахты определён по её вертикальным границам; проверьте технические уровни и уровни перекрытий.");
                            }
                            else r.Items.Add(item);
                        }
                        catch (Exception ex) { r.Error(s.Name + " / ID " + IDHelper.ElIdValue(e.Id) + ": " + ex.Message); }
                    }
                if (gns) ApplyVerticalRules(r);
                if (m.Index == 8)
                {
                    foreach (var a in r.Items.Where(x => !x.Excluded).GroupBy(x => x.Apartment)) r.ApartmentTotals[a.Key] = 1;
                    r.Value = r.ApartmentTotals.Count;
                    foreach (var i in r.Items.Where(i => i.Shape != null)) r.Pieces.Add(ToPiece(i));
                }
                else BuildPieces(r);
                if (!r.Items.Any(i => !i.Excluded)) r.Review("Не найдено ни одного включённого элемента; проверьте фильтры, стадию и параметры.");
                if (gns && (m.Primary == "Помещения" || m.Primary == "Пространства"))
                    r.Review("ГНС восстановлена по пространствам до осей стен с добавлением половины толщины наружных стен. Проверьте полноту помещений, наружную функцию стен, фасадные слои, шахты и границы частей. Незамоделированные пустоты не достраиваются.");
                if (gns && m.Primary == "Зоны") r.Review("Проверьте границы исходных зон: наружный контур стен, а между функциональными частями - ось стены.");
                if (gns && m.Primary == "Элементы") r.Review("Проверьте совпадение проекций выбранных элементов с наружным обмером стен.");
                if ((m.Index == 5 || m.Index == 6) && !m.Rules.Any(x => x.Role == "Летнее помещение"))
                    r.Review("Правила летних помещений не заданы. Проверьте, что летние помещения распознаны фильтрами включения/исключения.");
                if (m.Index != 8)
                {
                    var pieces = r.Pieces.Where(p => !p.Excluded).ToList();
                    for (int a = 0; a < pieces.Count; a++)
                        for (int b = a + 1; b < pieces.Count; b++)
                        {
                            var p = pieces[a]; var q = pieces[b];
                            if (p.Building == q.Building || Math.Abs(p.Elevation - q.Elevation) > .003) continue;
                            try
                            {
                                double overlap = AreaM2(BooleanOperationsUtils.ExecuteBooleanOperation(p.Shape, q.Shape, BooleanOperationsType.Intersect));
                                if (overlap > .01) r.Error("Корпуса «" + p.Building + "» и «" + q.Building + "» пересекаются на " + overlap.ToString("N3") + " м². Проверьте дубли связей или распределение общей части; их вклады пока сохранены раздельно.");
                            }
                            catch (Exception ex) { r.Review("Не удалось проверить пересечение корпусов: " + ex.Message); }
                        }
                }
            }
            private string Building(Item i, ResultRow r)
            {
                if (Settings.Grouping == "По связям Revit") return i.Source.Name;
                if (Settings.Grouping == "Ручной выбор")
                {
                    if (!Settings.SelectedOnly) { r.Error("Ручной выбор корпуса требует флажка «Только выбранные до запуска»."); return null; }
                    if (string.IsNullOrWhiteSpace(Settings.ManualBuilding)) { r.Error("Не задано имя корпуса."); return null; }
                    return Settings.ManualBuilding.Trim();
                }
                if (Settings.Grouping == "По рабочему набору")
                {
                    if (!i.Source.Document.IsWorkshared) { r.Error(i.Source.Name + ": модель не содержит рабочих наборов."); return null; }
                    return i.Source.Document.GetWorksetTable().GetWorkset(i.Element.WorksetId).Name;
                }
                string issue; string value = Read(i.Element, Settings.BuildingParameter, out issue);
                if (issue != null || string.IsNullOrWhiteSpace(value)) { r.Error(i.Description + ": не определён корпус; " + issue); return null; }
                return value.Trim();
            }
            private void ApplyVerticalRules(ResultRow r)
            {
                var groups = new Dictionary<string, List<Item>>();
                foreach (var i in r.Items.Where(i => !i.Excluded))
                {
                    bool once = CountOnOneFloor(Settings.Method, i.Role, i.Area);
                    if (!once) continue;
                    string issue = null;
                    var id = i.Element is Opening && IDHelper.ElIdValue(i.Element.Category.Id) == (long)BuiltInCategory.OST_ShaftOpening
                        ? i.Element.UniqueId : Read(i.Element, r.Metric.VoidParameter, out issue);
                    if (issue != null || string.IsNullOrWhiteSpace(id))
                    { i.Excluded = true; i.Reason = "Не задан ID вертикального пространства"; r.Error(i.Description + ": для учёта на одном этаже нужен ID вертикального пространства."); continue; }
                    string key = Key(i.Source.Key, i.Building, i.Role, id);
                    if (!groups.ContainsKey(key)) groups[key] = new List<Item>(); groups[key].Add(i);
                }
                foreach (var group in groups.Values)
                {
                    double low = group.Min(i => i.Elevation);
                    foreach (var i in group.Where(i => i.Elevation > low + .003)) { i.Excluded = true; i.Reason = "Учтено на нижнем этаже"; }
                    if (group.Count == 1) r.Review(group[0].Description + ": найден только один этаж вертикального пространства; проверьте наличие масок выше.");
                }
            }
            private void BuildPieces(ResultRow r)
            {
                bool gns = r.Metric.Index < 5;
                foreach (var i in r.Items.Where(i => i.Excluded && i.Shape != null)) r.Pieces.Add(ToPiece(i));
                foreach (var group in r.Items.Where(i => !i.Excluded && i.Shape != null).GroupBy(i => Key(i.Building, Math.Round(i.Elevation, 6).ToString(CultureInfo.InvariantCulture))))
                {
                    Solid occupied = null;
                    var masks = r.Items.Where(i => i.Excluded && i.Shape != null && i.Source.Mode != "Только проверка" &&
                        i.Building == group.First().Building && Math.Abs(i.Elevation - group.First().Elevation) < .003).ToList();
                    foreach (var item in group.OrderBy(i => i.Identity, StringComparer.Ordinal))
                    {
                        try
                        {
                            Solid shape = item.Shape;
                            foreach (var mask in masks) shape = Difference(shape, mask.Shape);
                            if (shape == null || shape.Volume < 1e-9) continue;
                            if (occupied != null)
                            {
                                double overlap = AreaM2(BooleanOperationsUtils.ExecuteBooleanOperation(shape, occupied, BooleanOperationsType.Intersect));
                                if (overlap > .001)
                                {
                                    if (!gns) { r.Error(item.Description + ": пересечение " + overlap.ToString("N3") + " м²; вклад пропущен, устраните дубли/пересечения."); continue; }
                                    r.Review(item.Description + ": пересечение " + overlap.ToString("N3") + " м² учтено один раз.");
                                }
                                // Remove every overlap, including those below the diagnostic tolerance.
                                shape = Difference(shape, occupied);
                            }
                            if (shape == null || shape.Volume < 1e-9) continue;
                            occupied = occupied == null ? shape : BooleanOperationsUtils.ExecuteBooleanOperation(occupied, shape, BooleanOperationsType.Union);
                            var piece = ToPiece(item); piece.Shape = shape; r.Pieces.Add(piece);
                            double area = AreaM2(shape) * item.Factor; r.Value += area;
                            if (!string.IsNullOrEmpty(item.Apartment))
                            { if (!r.ApartmentTotals.ContainsKey(item.Apartment)) r.ApartmentTotals[item.Apartment] = 0; r.ApartmentTotals[item.Apartment] += area; }
                        }
                        catch (Exception ex) { r.Error(item.Description + ": не удалось объединить/вычесть контуры; вклад пропущен: " + ex.Message); }
                    }
                }
            }
            private static Piece ToPiece(Item i)
            { return new Piece { Shape = i.Shape, Elevation = i.Elevation, Building = i.Building, Source = i.Source.Name, LevelName = i.LevelName, Factor = i.Factor, Excluded = i.Excluded, Description = i.Description + (i.Excluded ? " | " + i.Reason : "") }; }
            private static Solid Difference(Solid a, Solid b)
            { return a == null || a.Volume < 1e-9 ? null : BooleanOperationsUtils.ExecuteBooleanOperation(a, b, BooleanOperationsType.Difference); }
            private static void Reconcile(Run run, int total, int a, int b)
            {
                var t = run.Results.FirstOrDefault(x => x.Metric.Index == total);
                var x1 = run.Results.FirstOrDefault(x => x.Metric.Index == a); var x2 = run.Results.FirstOrDefault(x => x.Metric.Index == b);
                if (t == null || x1 == null || x2 == null) return;
                if (Math.Abs(t.Value - x1.Value - x2.Value) > .01)
                {
                    string msg = "Баланс «" + t.Metric.Title + "» - сумма частей = " + (t.Value - x1.Value - x2.Value).ToString("N3") + " м². Проверьте настройки и классификацию.";
                    t.Error(msg); x1.Error(msg); x2.Error(msg);
                }
                try
                {
                    foreach (var p in x1.Pieces.Where(p => !p.Excluded))
                        foreach (var q in x2.Pieces.Where(q => !q.Excluded && Math.Abs(q.Elevation - p.Elevation) < .003))
                            if (AreaM2(BooleanOperationsUtils.ExecuteBooleanOperation(p.Shape, q.Shape, BooleanOperationsType.Intersect)) > .01)
                            { x1.Error("Пересечение с «" + x2.Metric.Title + "»."); x2.Error("Пересечение с «" + x1.Metric.Title + "»."); return; }
                }
                catch (Exception ex) { t.Review("Не удалось проверить пересечения частей: " + ex.Message); }
            }
            // Normalized solids have height 1 internal foot at z=0: Volume equals footprint area.
            // The same solids supply numeric results, filled regions and 3D verification.
            private static Solid Footprint(Item i, bool gross, ResultRow r)
            {
                var opening = i.Element as Opening;
                if (opening != null)
                {
                    if (opening.IsRectBoundary || opening.BoundaryCurves == null)
                        throw new InvalidOperationException("У проёма нет горизонтального замкнутого контура; задайте поэтажную зону.");
                    return Extrude(new List<CurveLoop> { Flatten(opening.BoundaryCurves.Cast<Curve>(), i.Source.Transform) });
                }
                var spatial = i.Element as SpatialElement;
                if (spatial != null)
                {
                    if (SpatialArea(spatial) <= 1e-9) throw new InvalidOperationException("Пространственный объект не размещён, не замкнут или имеет нулевую площадь.");
                    using (var options = new SpatialElementBoundaryOptions())
                    {
                        options.SpatialElementBoundaryLocation = gross && !(spatial is Area) ? SpatialElementBoundaryLocation.Center : SpatialElementBoundaryLocation.Finish;
                        var bounds = spatial.GetBoundarySegments(options);
                        if (bounds == null || bounds.Count == 0) throw new InvalidOperationException("Не найдены замкнутые границы.");
                        var loops = bounds.Select(b => Flatten(b.Select(s => s.GetCurve()), i.Source.Transform)).ToList();
                        Solid centerShape = Extrude(loops);
                        if (gross && !(spatial is Area))
                        {
                            for (int loopIndex = 0; loopIndex < bounds.Count; loopIndex++)
                            {
                                var distances = new List<double>();
                                foreach (var segment in bounds[loopIndex])
                                {
                                    double distance = 0;
                                    var boundaryElement = i.Element.Document.GetElement(segment.ElementId);
                                    var wall = boundaryElement as Wall;
                                    if (boundaryElement is RevitLinkInstance)
                                        throw new InvalidOperationException("Граница по стене другой связи: используйте зону по внешнему контуру.");
                                    if (wall != null && wall.WallType.Function == WallFunction.Exterior)
                                    {
                                        if (wall.WallType.Kind != WallKind.Basic)
                                            throw new InvalidOperationException("Наружная стена не базового типа: используйте зону по внешнему контуру.");
                                        var curve = segment.GetCurve().CreateTransformed(i.Source.Transform);
                                        XYZ tangent = curve.ComputeDerivatives(.5, true).BasisX.Normalize();
                                        XYZ right = tangent.CrossProduct(XYZ.BasisZ).Normalize();
                                        XYZ point = curve.Evaluate(.5, true) + right * .001;
                                        using (var opt = new SolidCurveIntersectionOptions())
                                        using (var intersection = centerShape.IntersectWithCurve(Line.CreateBound(new XYZ(point.X, point.Y, .25), new XYZ(point.X, point.Y, .75)), opt))
                                            distance = (intersection.SegmentCount > 0 ? -1 : 1) * wall.Width / 2;
                                    }
                                    distances.Add(distance);
                                }
                                // Variable offsets preserve axes of internal partition walls and trim corners exactly.
                                if (distances.Any(d => Math.Abs(d) > 1e-9))
                                    loops[loopIndex] = CurveLoop.CreateViaOffset(loops[loopIndex], distances, XYZ.BasisZ);
                            }
                        }
                        Solid shape = Extrude(loops);
                        if (!gross && Math.Abs(AreaM2(shape) - IDHelper.ConvertInternalAreaToSquareMeters(SpatialArea(spatial))) > .02)
                            r.Review(i.Description + ": площадь контура отличается от параметра Revit; используется проверяемый контур.");
                        return shape;
                    }
                }
                var solids = new List<Solid>();
                using (var options = new Options { DetailLevel = ViewDetailLevel.Fine, IncludeNonVisibleObjects = false }) CollectSolids(i.Element.get_Geometry(options), solids);
                Solid result = null;
                foreach (var solid in solids)
                    foreach (Face face in solid.Faces)
                    {
                        var plane = face as PlanarFace;
                        if (plane == null || i.Source.Transform.OfVector(plane.FaceNormal).Z <= 1e-6) continue;
                        var loops = plane.GetEdgesAsCurveLoops().Select(loop => Flatten(loop, i.Source.Transform)).ToList();
                        var part = Extrude(loops);
                        result = result == null ? part : BooleanOperationsUtils.ExecuteBooleanOperation(result, part, BooleanOperationsType.Union);
                    }
                if (result == null || result.Volume < 1e-9) throw new InvalidOperationException("Нет плоских верхних граней для проекции. Используйте помещение/зону с точным контуром.");
                return result;
            }
            private static void CollectSolids(GeometryElement geometry, List<Solid> result)
            {
                if (geometry == null) return;
                foreach (GeometryObject obj in geometry)
                {
                    var solid = obj as Solid; if (solid != null && solid.Volume > 1e-9) result.Add(solid);
                    var instance = obj as GeometryInstance; if (instance != null) CollectSolids(instance.GetInstanceGeometry(), result);
                }
            }
            private static CurveLoop Flatten(IEnumerable<Curve> curves, Transform transform)
            {
                var result = new CurveLoop();
                foreach (var raw in curves)
                {
                    var c = raw.CreateTransformed(transform);
                    if (c is Line)
                    {
                        var a = c.GetEndPoint(0); var b = c.GetEndPoint(1); a = new XYZ(a.X, a.Y, 0); b = new XYZ(b.X, b.Y, 0);
                        if (a.DistanceTo(b) > 1e-6) result.Append(Line.CreateBound(a, b));
                    }
                    else if (c is Arc && Math.Abs(((Arc)c).Normal.Z) > .999999)
                        result.Append(c.CreateTransformed(Transform.CreateTranslation(new XYZ(0, 0, -c.GetEndPoint(0).Z))));
                    else throw new InvalidOperationException("Сложная кривая не проецируется без аппроксимации; задайте точную зону.");
                }
                if (result.IsOpen()) throw new InvalidOperationException("Незамкнутый контур."); return result;
            }
            private static Solid Extrude(IList<CurveLoop> loops)
            { return GeometryCreationUtilities.CreateExtrusionGeometry(loops, XYZ.BasisZ, 1.0); }
            public string GroundFromSelection()
            {
                var heights = new List<double>();
                foreach (var s in Sources.Where(s => s.Loaded))
                    foreach (var e in s.Elements.Where(e => (s.Key == "host" && selection.Contains(IDHelper.ElIdValue(e.Id))) ||
                        (s.Key != "host" && selection.Contains(IDHelper.ElIdValue(s.RootLink)))))
                    {
                        var topo = e as TopographySurface;
                        if (topo != null) heights.AddRange(topo.GetPoints().Select(p => s.Transform.OfPoint(p).Z));
                        else if (e.GetType().Name == "Toposolid")
                        {
                            var solids = new List<Solid>(); using (var op = new Options()) CollectSolids(e.get_Geometry(op), solids);
                            foreach (var solid in solids) foreach (Face face in solid.Faces)
                                {
                                    var uv = face.GetBoundingBox();
                                    if (face.ComputeNormal((uv.Min + uv.Max) * .5).Z > 0) heights.AddRange(face.Triangulate().Vertices.Select(p => s.Transform.OfPoint(p).Z));
                                }
                        }
                    }
                if (heights.Count == 0) return "Среди выбранных до запуска элементов/связей нет топографии/топотела. Введите отметку земли вручную.";
                Settings.Ground = (heights.Average() * .3048).ToString("0.000");
                return "Справочная отметка: " + Settings.Ground + " м - среднее высот вершин поверхности, не средняя планировочная отметка по периметру здания. Проверьте её и выберите первый этаж с учётом правила +2 м для цоколя.";
            }
            private static Schema StorageSchema()
            {
                var schema = Schema.Lookup(SchemaId); if (schema != null) return schema;
                var b = new SchemaBuilder(SchemaId); b.SetSchemaName("KPLNCalculateTEP_v1"); b.SetReadAccessLevel(AccessLevel.Public); b.SetWriteAccessLevel(AccessLevel.Public);
                b.AddSimpleField("Kind", typeof(string)); b.AddSimpleField("Payload", typeof(string)); return b.Finish();
            }
            private static void Tag(Element e, string kind, string payload)
            {
                var s = StorageSchema(); var entity = new Entity(s); entity.Set(s.GetField("Kind"), kind); entity.Set(s.GetField("Payload"), payload); e.SetEntity(entity);
            }
            private static string Tagged(Element e, string field)
            {
                var s = Schema.Lookup(SchemaId); if (s == null) return null;
                var entity = e.GetEntity(s); return entity.IsValid() ? entity.Get<string>(s.GetField(field)) : null;
            }
            public void SaveSettings()
            {
                using (var tx = new Transaction(doc, "ТЭП: сохранить настройки"))
                {
                    tx.Start();
                    var failures = new List<string>();
                    tx.SetFailureHandlingOptions(tx.GetFailureHandlingOptions().SetFailuresPreprocessor(new FailureLog(failures)).SetClearAfterRollback(true));
                    var storage = new FilteredElementCollector(doc).OfClass(typeof(DataStorage)).FirstOrDefault(e => Tagged(e, "Kind") == "settings") ?? DataStorage.Create(doc);
                    Tag(storage, "settings", Serialize().ToString());
                    if (tx.Commit() != TransactionStatus.Committed) throw new InvalidOperationException("Транзакция настроек отменена. " + string.Join("; ", failures));
                }
            }
            private XElement Serialize()
            {
                var root = new XElement("TEP", new XAttribute("version", "1"));
                foreach (string p in new[] { "Method", "Grouping", "BuildingParameter", "ManualBuilding", "SelectedOnly", "Ground" })
                    root.SetAttributeValue(p, typeof(Settings).GetProperty(p).GetValue(Settings, null) ?? "");
                foreach (var pair in new[] { new { Name = "Old", List = Settings.Old }, new { Name = "New", List = Settings.New } })
                {
                    var list = new XElement(pair.Name); root.Add(list);
                    foreach (var metric in pair.List)
                    {
                        var node = new XElement("Metric"); list.Add(node);
                        foreach (var p in typeof(Metric).GetProperties().Where(p => p.CanWrite && p.Name != "Rules")) node.SetAttributeValue(p.Name, p.GetValue(metric, null) ?? "");
                        foreach (var rule in metric.Rules)
                        { var n = new XElement("Rule"); node.Add(n); foreach (var p in typeof(Rule).GetProperties().Where(p => p.CanWrite)) n.SetAttributeValue(p.Name, p.GetValue(rule, null) ?? ""); }
                    }
                }
                foreach (var source in Sources) root.Add(new XElement("Source", new XAttribute("key", source.Key), new XAttribute("mode", source.Mode)));
                return root;
            }
            private Settings LoadSettings()
            {
                var settings = new Settings();
                try
                {
                    var storage = new FilteredElementCollector(doc).OfClass(typeof(DataStorage)).FirstOrDefault(e => Tagged(e, "Kind") == "settings");
                    if (storage == null) return settings;
                    var root = XElement.Parse(Tagged(storage, "Payload"));
                    if ((string)root.Attribute("version") != "1") throw new InvalidOperationException("Неизвестная версия настроек.");
                    ReadProperties(settings, root);
                    foreach (var pair in new[] { new { Name = "Old", List = settings.Old }, new { Name = "New", List = settings.New } })
                        foreach (var node in root.Element(pair.Name).Elements("Metric"))
                        {
                            var m = new Metric(); ReadProperties(m, node); if (m.Index < 0 || m.Index > 8) continue;
                            foreach (var n in node.Elements("Rule")) { var rule = new Rule(); ReadProperties(rule, n); m.Rules.Add(rule); }
                            pair.List[m.Index] = m;
                        }
                    foreach (var n in root.Elements("Source")) { var s = Sources.FirstOrDefault(x => x.Key == (string)n.Attribute("key")); if (s != null) s.Mode = (string)n.Attribute("mode"); }
                }
                catch (Exception ex) { scanMessages.Add("Настройки не загружены: " + ex.Message); settings = new Settings(); }
                return settings;
            }
            private static void ReadProperties(object obj, XElement node)
            {
                foreach (var p in obj.GetType().GetProperties().Where(p => p.CanWrite))
                {
                    var a = node.Attribute(p.Name); if (a == null) continue;
                    if (p.PropertyType == typeof(string)) p.SetValue(obj, a.Value, null);
                    else if (p.PropertyType == typeof(bool)) p.SetValue(obj, bool.Parse(a.Value), null);
                    else if (p.PropertyType == typeof(int)) p.SetValue(obj, int.Parse(a.Value, CultureInfo.InvariantCulture), null);
                    else if (p.PropertyType == typeof(double)) p.SetValue(obj, double.Parse(a.Value, CultureInfo.InvariantCulture), null);
                }
            }
            public string CreateGraphics()
            {
                if (LastRun == null) throw new InvalidOperationException("Сначала выполните расчёт.");
                if (!LastRun.Results.Any(r => r.Pieces.Any())) throw new InvalidOperationException("Нет рассчитанной геометрии.");
                int created = 0; var errors = new List<string>();
                using (var tx = new Transaction(doc, "ТЭП: обновить проверочные виды и контуры"))
                {
                    tx.Start();
                    tx.SetFailureHandlingOptions(tx.GetFailureHandlingOptions().SetFailuresPreprocessor(new FailureLog(errors)).SetClearAfterRollback(true));
                    // Only elements with our ExtensibleStorage ownership tag may be replaced.
                    foreach (var id in new FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
                        .Where(e => Tagged(e, "Kind") == "graphic").Select(e => e.Id).ToList()) doc.Delete(id);
                    var floorType = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault(v => v.ViewFamily == ViewFamily.FloorPlan);
                    var threeType = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault(v => v.ViewFamily == ViewFamily.ThreeDimensional);
                    if (floorType == null || threeType == null) throw new InvalidOperationException("Нет типов плана этажа или 3D-вида.");
                    var fill = new FilteredElementCollector(doc).OfClass(typeof(FilledRegionType)).Cast<FilledRegionType>().FirstOrDefault(t => Tagged(t, "Kind") == "filltype");
                    if (fill == null)
                    {
                        var seed = new FilteredElementCollector(doc).OfClass(typeof(FilledRegionType)).Cast<FilledRegionType>().FirstOrDefault();
                        if (seed == null) throw new InvalidOperationException("Загрузите в проект хотя бы один тип цветовой области.");
                        fill = (FilledRegionType)seed.Duplicate("ТЭП заливка " + Guid.NewGuid().ToString("N").Substring(0, 6)); Tag(fill, "filltype", "");
                    }
                    var pattern = new FilteredElementCollector(doc).OfClass(typeof(FillPatternElement)).Cast<FillPatternElement>().FirstOrDefault(f => f.GetFillPattern().IsSolidFill);
                    if (pattern == null) throw new InvalidOperationException("Нет сплошной штриховки.");
                    fill.ForegroundPatternId = pattern.Id; fill.IsMasking = false;
                    var allShapes = new List<ElementId>(); var owners = new Dictionary<View3D, List<ElementId>>();
                    foreach (var result in LastRun.Results.Where(r => r.Pieces.Count > 0))
                    {
                        string viewKey = "3d/" + result.Metric.Index;
                        var view3 = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().FirstOrDefault(v => Tagged(v, "Payload") == viewKey && Tagged(v, "Kind") == "view");
                        if (view3 == null) { view3 = View3D.CreateIsometric(doc, threeType.Id); view3.Name = UniqueViewName("ТЭП · " + result.Metric.Index + " · 3D"); Tag(view3, "view", viewKey); }
                        owners[view3] = new List<ElementId>();
                        foreach (var p in result.Pieces)
                        {
                            using (var sub = new SubTransaction(doc))
                            {
                                sub.Start();
                                try
                                {
                                    var level = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().FirstOrDefault(l => Math.Abs(l.ProjectElevation - p.Elevation) < .003);
                                    if (level == null) { level = Level.Create(doc, p.Elevation); level.Name = "ТЭП проверка " + (p.Elevation * .3048).ToString("0.000") + " " + Guid.NewGuid().ToString("N").Substring(0, 4); Tag(level, "level", ""); }
                                    string planKey = "plan/" + result.Metric.Index + "/" + level.UniqueId;
                                    var plan = new FilteredElementCollector(doc).OfClass(typeof(ViewPlan)).Cast<ViewPlan>().FirstOrDefault(v => Tagged(v, "Payload") == planKey && Tagged(v, "Kind") == "view");
                                    if (plan == null)
                                    {
                                        plan = ViewPlan.Create(doc, floorType.Id, level.Id); plan.Name = UniqueViewName("ТЭП · " + result.Metric.Index + " · " + level.Name);
                                        Tag(plan, "view", planKey); plan.CropBoxActive = false;
                                    }
                                    var phase = doc.Phases.Cast<Phase>().FirstOrDefault(f => f.Name == result.Metric.Phase);
                                    if (phase != null)
                                    {
                                        var pp = plan.get_Parameter(BuiltInParameter.VIEW_PHASE); if (pp != null && !pp.IsReadOnly) pp.Set(phase.Id);
                                        pp = view3.get_Parameter(BuiltInParameter.VIEW_PHASE); if (pp != null && !pp.IsReadOnly) pp.Set(phase.Id);
                                    }
                                    Color color = p.Excluded ? new Color(205, 65, 65) : result.Metric.Index == 5 || result.Metric.Index == 6 ? new Color(55, 150, 95) : new Color(50, 120, 210);
                                    var ov = new OverrideGraphicSettings().SetSurfaceForegroundPatternId(pattern.Id).SetSurfaceForegroundPatternColor(color).SetProjectionLineColor(color).SetSurfaceTransparency(40);
                                    foreach (Face f in p.Shape.Faces)
                                    {
                                        var pf = f as PlanarFace; if (pf == null || pf.FaceNormal.Z < .999999) continue;
                                        var loops = pf.GetEdgesAsCurveLoops().Select(l => CurveLoop.CreateViaTransform(l, Transform.CreateTranslation(new XYZ(0, 0, level.ProjectElevation - 1)))).ToList();
                                        var region = FilledRegion.Create(doc, fill.Id, plan.Id, loops); Tag(region, "graphic", p.Description); plan.SetElementOverrides(region.Id, ov);
                                    }
                                    var ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
                                    ds.ApplicationId = AppId; ds.ApplicationDataId = result.Metric.Index + "/" + Guid.NewGuid().ToString("N");
                                    ds.Name = "ТЭП " + (p.Excluded ? "Исключено" : "Учтено") + " · " + result.Metric.Title;
                                    ds.SetShape(new List<GeometryObject> { SolidUtils.CreateTransformed(p.Shape, Transform.CreateTranslation(new XYZ(0, 0, p.Elevation))) });
                                    Tag(ds, "graphic", p.Description);
                                    var comment = ds.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                                    if (comment != null && !comment.IsReadOnly) comment.Set(p.Description + " | " + AreaM2(p.Shape).ToString("N3") + " м² × " + p.Factor.ToString("0.###"));
                                    view3.SetElementOverrides(ds.Id, ov);
                                    if (sub.Commit() != TransactionStatus.Committed) throw new InvalidOperationException("Подтранзакция графики отменена.");
                                    allShapes.Add(ds.Id); owners[view3].Add(ds.Id); created++;
                                }
                                catch (Exception ex)
                                {
                                    if (sub.GetStatus() == TransactionStatus.Started) sub.RollBack();
                                    errors.Add(p.Description + ": " + ex.Message);
                                }
                            }
                        }
                    }
                    if (created == 0) throw new InvalidOperationException("Не создан ни один контур; прежняя графика сохранена. " + string.Join("; ", errors));
                    // Hide verification slabs in existing ordinary views and other metrics' views.
                    foreach (var view in new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>().Where(v => !v.IsTemplate))
                    {
                        var v3 = view as View3D;
                        var own = v3 != null && owners.ContainsKey(v3) ? owners[v3] : new List<ElementId>();
                        var hidden = allShapes.Where(id => !own.Contains(id) && doc.GetElement(id).CanBeHidden(view)).ToList();
                        if (hidden.Count > 0) view.HideElements(hidden);
                    }
                    if (tx.Commit() != TransactionStatus.Committed) throw new InvalidOperationException("Проверочная графика не сохранена; прежняя версия не заменена. " + string.Join("; ", errors));
                }
                LastRun.GraphicsStatus = "Проверочная графика: " + created + " фрагментов. Синий - учтённая площадь; зелёный - квартиры; красный - исключённая/справочная площадь.\n" +
                    "3D-пластины высотой 0,3048 м служат проверке площади, это не строительный объём. Виды начинаются с «ТЭП».\n" + string.Join("\n", errors.Select(e => "ОШИБКА ГРАФИКИ: " + e));
                return LastRun.GraphicsStatus;
            }
            private string UniqueViewName(string name)
            {
                var names = new HashSet<string>(new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>().Select(v => v.Name));
                string value = name; int suffix = 2; while (names.Contains(value)) value = name + " (" + suffix++ + ")"; return value;
            }
        }
        private static double SpatialArea(Element e)
        { var spatial = e as SpatialElement; if (spatial != null) return spatial.Area; var p = e.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED); return p == null ? 0 : p.AsDouble(); }
        public static double AreaM2(Solid shape) { return shape == null ? 0 : IDHelper.ConvertInternalAreaToSquareMeters(shape.Volume); }
        public static bool TryNumber(string value, out double result)
        { return double.TryParse((value ?? "").Trim().Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out result) && !double.IsNaN(result) && !double.IsInfinity(result); }
        public static bool CountOnOneFloor(string method, string role, double area)
        {
            return role == "Многосветное пространство" || role == "Лестничный просвет" ||
                role == "Проём" && (method == "Новая" || area > 36.0) || role == "Шахта" && method == "Новая";
        }
        private static bool ReadFactor(Element e, string name, out double value)
        {
            value = 0;
            var ps = e.GetParameters(name).ToList();
            if (ps.Count == 0) { var type = e.Document.GetElement(e.GetTypeId()); if (type != null) ps = type.GetParameters(name).ToList(); }
            if (ps.Count != 1 || !ps[0].HasValue) return false;
            var p = ps[0];
            if (p.StorageType == StorageType.Double && IDHelper.IsNumber(p)) { value = p.AsDouble(); return !double.IsNaN(value) && !double.IsInfinity(value); }
            if (p.StorageType == StorageType.Integer) { value = p.AsInteger(); return true; }
            return p.StorageType == StorageType.String && TryNumber(p.AsString(), out value);
        }
        private static string Key(params string[] parts)
        { return string.Join(" | ", parts.Select(p => { var s = (p ?? "").Trim().ToUpperInvariant(); return s.Length + ":" + s; })); }
        private static string Read(Element e, string name, out string issue)
        {
            issue = null;
            if (string.IsNullOrWhiteSpace(name)) { issue = "Не задано имя параметра."; return null; }
            var candidates = e.GetParameters(name).ToList();
            if (candidates.Count == 0) { var type = e.Document.GetElement(e.GetTypeId()); if (type != null) candidates = type.GetParameters(name).ToList(); }
            if (candidates.Count != 1) { issue = candidates.Count == 0 ? "Нет параметра «" + name + "»." : "Неоднозначное имя параметра «" + name + "»."; return null; }
            var p = candidates[0]; if (!p.HasValue) return "";
            if (p.StorageType == StorageType.String) return (p.AsString() ?? "").Trim();
            if (p.StorageType == StorageType.Integer) return p.AsInteger().ToString(CultureInfo.InvariantCulture);
            if (p.StorageType == StorageType.Double) return p.AsValueString() ?? p.AsDouble().ToString(CultureInfo.InvariantCulture);
            if (p.StorageType == StorageType.ElementId) { var referenced = e.Document.GetElement(p.AsElementId()); return referenced == null ? "" : referenced.Name; }
            return "";
        }
        private static bool Matches(Element e, Rule r, out string issue)
        {
            issue = null;
            if (!string.IsNullOrWhiteSpace(r.Category) && e.Category.Name != r.Category) return false;
            if (string.IsNullOrWhiteSpace(r.Parameter)) return !string.IsNullOrWhiteSpace(r.Category);
            string value = Read(e, r.Parameter, out issue); if (issue != null) return false;
            if (r.Operation == "Не пусто") return !string.IsNullOrWhiteSpace(value);
            if (r.Operation == "Пусто") return string.IsNullOrWhiteSpace(value);
            if (r.Operation == "Содержит") return !string.IsNullOrEmpty(r.Value) && (value ?? "").IndexOf(r.Value.Trim(), StringComparison.OrdinalIgnoreCase) >= 0;
            return string.Equals(value, (r.Value ?? "").Trim(), StringComparison.OrdinalIgnoreCase);
        }
    }
}