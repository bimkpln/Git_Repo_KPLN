using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Tools_OVVK.Forms;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace KPLN_Tools_OVVK.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Command_VentilationSettingsConfigurator : IExternalCommand
    {
        internal const string PluginName = "Конфигуратор вент. установок";
        internal const string UnknownProjectName = "Неизвестно";
        internal const string ProjectDatabasePath =
            @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_Loader_MainDB.db";
        internal const string ProjectFamiliesRoot =
            @"Z:\Отдел BIM\07_Вспомогательное\Семейства\03_Вентустановки";
        internal const string ProjectFamilyFileName = "550_Универсальная установка_Одноуровневая_(Об).rfa";
        internal const string BaseFamilyTypeName =
            "имяСистемы(при необходимости)_имяУстановкиПоПодборке";
        // Известный демонстрационный тип исходного семейства сохраняется без изменений.
        internal const string ConfiguredFamilyTypeName =
            "П6_KT 2,2-250911903-02.02-K-O-P-A";

        // Обычный путь Windows: обратная косая черта перед '_' в сообщении
        // может быть экранированием Markdown, а не разделителем папок.
        internal const string SourceFamilyPath =
            @"X:\BIM\3_Семейства\4_ОВиК\4_ОВ2_Вентиляция\550-596_Инженерное оборудование\550_Универсальная установка_Одноуровневая_(Об).rfa";
        private const string LiteralSourceFamilyPath =
            @"X:\BIM\3\_Семейства\4\_ОВиК\4\_ОВ2\_Вентиляция\550-596\_Инженерное оборудование\550\_Универсальная установка\_Одноуровневая\_(Об).rfa";

        private static VentilationSettingsConfiguratorMain _window;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                if (_window == null || !_window.IsLoaded)
                {
                    var uiapp = commandData.Application;
                    var window = new VentilationSettingsConfiguratorMain(uiapp, uiapp.ActiveUIDocument);
                    _window = window;
                    window.Closed += (s, e) => { if (ReferenceEquals(_window, window)) _window = null; };
                    window.Show();
                }
                else
                {
                    if (_window.WindowState == System.Windows.WindowState.Minimized)
                        _window.WindowState = System.Windows.WindowState.Normal;
                    _window.Activate();
                }
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                if (_window != null && !_window.IsLoaded)
                {
                    _window.ReleaseExternalEvent();
                    _window = null;
                }
                message = ex.Message;
                return Result.Failed;
            }
        }

        internal static string ResolveSourcePath()
        {
            // Сначала проверяем буквальный вариант, затем вариант без Markdown-экранирования.
            return File.Exists(LiteralSourceFamilyPath) ? LiteralSourceFamilyPath : SourceFamilyPath;
        }

        internal enum RequestKind { LoadSectionCatalog, Save, OpenInRevit, AddToProject, DeleteType }

        // CONFIGURATION MODEL BEGIN — только данные; ссылки на Revit Document / Element здесь не храним.
        internal abstract class EditableValue : INotifyPropertyChanged
        {
            public event PropertyChangedEventHandler PropertyChanged;
            protected void Changed([CallerMemberName] string name = null)
            { PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name)); }
        }

        internal enum ConnectionKind { None, WaterHeater, WaterCooler, RefrigerantCooler }

        internal static class SharedParameters
        {
            internal const string FrameVisible = "Рама_Видимость", ServiceRight = "ЗО_Справа", ServiceDepth = "ЗО_Глубина";
            internal const string HeaterRight = "Нагреватель_Трубки_Справа", CoolerRight = "Охладитель_Трубки_Справа";
            internal static readonly string[] WaterLengths = {
                "Вода_Патрубок_Обратка_Смещение X", "Вода_Патрубок_Обратка_Смещение Y",
                "Вода_Патрубок_Приток/Обратка_Диаметр", "Вода_Патрубок_Приток_Смещение X", "Вода_Патрубок_Приток_Смещение Y" };
            internal static readonly string[] RefrigerantLengths = {
                "Дренаж_Патрубок_Диаметр", "Дренаж_Патрубок_Смещение X", "Дренаж_Патрубок_Смещение Y",
                "Фреон_Патрубок_Обратка_Диаметр", "Фреон_Патрубок_Обратка_Смещение X", "Фреон_Патрубок_Обратка_Смещение Y",
                "Фреон_Патрубок_Приток_Диаметр", "Фреон_Патрубок_Приток_Смещение X", "Фреон_Патрубок_Приток_Смещение Y" };
            internal static IEnumerable<string> LengthNames { get { return new[] { ServiceDepth }.Concat(WaterLengths).Concat(RefrigerantLengths); } }
            internal static readonly string[] BooleanNames = { FrameVisible, ServiceRight, HeaterRight, CoolerRight };
            internal static IEnumerable<string> ForSection(SectionTypeChoice type)
            {
                var kind = type == null ? ConnectionKind.None : type.Connections;
                if (kind == ConnectionKind.WaterHeater || kind == ConnectionKind.WaterCooler)
                    return new[] { kind == ConnectionKind.WaterHeater ? HeaterRight : CoolerRight }.Concat(WaterLengths);
                if (kind == ConnectionKind.RefrigerantCooler) return new[] { CoolerRight }.Concat(RefrigerantLengths);
                return Enumerable.Empty<string>();
            }
            internal static string Label(string name)
            {
                if (name == FrameVisible) return "Рама";
                if (name == ServiceRight) return "Зона обслуживания";
                if (name == ServiceDepth) return "Зона обслуживания: глубина";
                if (name == HeaterRight || name == CoolerRight) return "Сторона трубок";
                return name.Replace("_Патрубок_", ": ").Replace('_', ' ');
            }
        }

        internal sealed class SectionTypeChoice
        {
            public string FamilyName { get; private set; }
            public string TypeName { get; private set; }
            public string DisplayName { get { return FamilyName + " : " + TypeName; } }
            internal bool IsEmpty { get { return string.Equals(TypeName, "Пустой блок", StringComparison.OrdinalIgnoreCase); } }
            internal bool IsFlexibleConnector
            {
                get
                {
                    return FamilyName.IndexOf("_Секция_Гибкая вставка_", StringComparison.OrdinalIgnoreCase) >= 0
                || string.Equals(TypeName, "Вставка гибкая", StringComparison.OrdinalIgnoreCase) || string.Equals(TypeName, "Гибкая вставка", StringComparison.OrdinalIgnoreCase);
                }
            }
            internal bool IsConnectorType
            {
                get
                {
                    return IsEmpty || IsFlexibleConnector || IsEquipment("Шумоглушитель") || FamilyName.IndexOf("_Секция_Воздушный клапан_", StringComparison.OrdinalIgnoreCase) >= 0
                || string.Equals(TypeName, "Воздушный клапан", StringComparison.OrdinalIgnoreCase);
                }
            }
            private bool IsEquipment(string name)
            {
                return FamilyName.IndexOf("_Секция_" + name + "_", StringComparison.OrdinalIgnoreCase) >= 0
                || string.Equals(TypeName, name, StringComparison.OrdinalIgnoreCase);
            }
            internal ConnectionKind Connections
            {
                get
                {
                    if (IsEquipment("Водяной нагреватель")) return ConnectionKind.WaterHeater;
                    if (IsEquipment("Водяной охладитель")) return ConnectionKind.WaterCooler;
                    if (IsEquipment("Фреоновый охладитель")) return ConnectionKind.RefrigerantCooler;
                    return ConnectionKind.None;
                }
            }
            internal string Key { get { return FamilyName + "\n" + TypeName; } }
            internal SectionTypeChoice(string familyName, string typeName)
            { FamilyName = familyName; TypeName = typeName; }
        }

        internal sealed class SectionDefinition
        {
            internal string ParameterName, DisplayName;
            internal List<SectionTypeChoice> Choices;
            internal SectionTypeChoice SourceType;
            internal double DefaultLengthMm, DefaultWidthMm, DefaultHeightMm, DefaultOffsetXMm, DefaultOffsetYMm;
            internal static List<SectionTypeChoice> OrderChoices(IEnumerable<SectionTypeChoice> choices)
            {
                return choices.GroupBy(c => c.Key).Select(g => g.First()).OrderBy(c => c.IsEmpty ? 0 : 1)
                    .ThenBy(c => c.DisplayName, StringComparer.CurrentCultureIgnoreCase).ToList();
            }
        }

        // Нумерация физического последнего блока остаётся 12. В другой редакции RFA его параметры могут иметь номер 2.
        // «Секция_2_Соединитель_Клапан_…» всегда обозначает отдельный второй блок, не последний соединитель.
        internal static class FamilyParameterNames
        {
            internal static string Canonical(string name)
            {
                string result = Regex.Replace((name ?? "").Replace("Соеденитель", "Соединитель"), @"Смещение_по_([XY])", "Смещение по $1");
                return Regex.Replace(result, @"Секция_2_Соединитель_(?!Клапан_)", "Секция_12_Соединитель_");
            }
            internal static IEnumerable<string> Aliases(string name)
            {
                string canonical = Canonical(name);
                foreach (string numbered in new[] { canonical, canonical.Replace("Секция_12_Соединитель_", "Секция_2_Соединитель_") }.Distinct())
                    foreach (string spelling in new[] { numbered, numbered.Replace("Соединитель", "Соеденитель") }.Distinct())
                        foreach (string spacing in new[] { spelling, spelling.Replace("Смещение по ", "Смещение_по_") }.Distinct())
                            yield return spacing;
            }
            internal static bool IsOffset(string name)
            {
                return Regex.IsMatch(Canonical(name), @"^Секция_(1|12)_Соединитель_Смещение по [XY]$")
                || Regex.IsMatch(name ?? "", @"^(Вода|Фреон|Дренаж)_Патрубок_(?:(?:Приток|Обратка)_)?Смещение [XY]$");
            }
        }

        internal sealed class SectionCatalog : Dictionary<int, SectionDefinition>
        {
            internal string SourceTypeName;
            internal double InstallationWidthMm, InstallationHeightMm, FrameHeightMm;
            internal int IntermediateCount;
            internal bool HasValve;
            internal Dictionary<string, double> SharedLengthsMm = new Dictionary<string, double>();
            internal Dictionary<string, bool?> SharedFlags = new Dictionary<string, bool?>();
            internal FamilyDependencies Dependencies = new FamilyDependencies();
            internal DimensionValue Dimension(string name, double millimeters)
            { return new DimensionValue(millimeters, Dependencies.IsCalculated(name), FamilyParameterNames.IsOffset(name)); }
        }

        internal sealed class ParameterDependency
        {
            internal string Name, Formula, SourceValue;
            internal int? IntegerValue;
            internal bool IsReadOnly, IsReporting, IsInstance;
            internal List<string> Inputs = new List<string>(), Dependents = new List<string>(), Links = new List<string>();
        }

        internal sealed class FamilyDependencies
        {
            internal Dictionary<string, ParameterDependency> Parameters = new Dictionary<string, ParameterDependency>(StringComparer.Ordinal);
            internal List<string> ReadWarnings = new List<string>();
            internal ParameterDependency Find(string name)
            {
                var matches = FamilyParameterNames.Aliases(name).Where(Parameters.ContainsKey).Select(n => Parameters[n]).ToList();
                if (matches.Count > 1) throw new InvalidOperationException("Неоднозначное соответствие параметра «" + name + "»: " + string.Join(", ", matches.Select(p => p.Name)));
                return matches.FirstOrDefault();
            }
            internal bool IsCalculated(string name)
            {
                var parameter = Find(name);
                return parameter != null
                    && (parameter.IsReadOnly || parameter.IsReporting || !string.IsNullOrWhiteSpace(parameter.Formula));
            }
            internal static IEnumerable<string> FormulaReferences(string formula, IEnumerable<string> names)
            {
                if (string.IsNullOrWhiteSpace(formula)) return Enumerable.Empty<string>();
                // Имена ищем целиком, сначала самые длинные; строковые константы не являются ссылками.
                string unquoted = Regex.Replace(formula, "\"(?:[^\"]|\"\")*\"", " ");
                string alternatives = string.Join("|", names.Where(n => !string.IsNullOrWhiteSpace(n)).Distinct()
                    .OrderByDescending(n => n.Length).Select(Regex.Escape));
                if (alternatives.Length == 0) return Enumerable.Empty<string>();
                return Regex.Matches(unquoted, @"(?<![\p{L}\p{N}_])(?:" + alternatives + @")(?![\p{L}\p{N}_])")
                    .Cast<Match>().Select(m => m.Value).Distinct().ToList();
            }
            internal void ConnectFormulas()
            {
                foreach (var parameter in Parameters.Values) { parameter.Inputs.Clear(); parameter.Dependents.Clear(); }
                foreach (var parameter in Parameters.Values)
                    foreach (string input in FormulaReferences(parameter.Formula, Parameters.Keys))
                    { parameter.Inputs.Add(input); Parameters[input].Dependents.Add(parameter.Name); }
            }
            internal string Tooltip(string name)
            {
                var parameter = Find(name);
                if (parameter == null) return name;
                var lines = new List<string> { parameter.Name };
                lines.Add(string.IsNullOrWhiteSpace(parameter.Formula) ? "Формула отсутствует." : "Формула: " + parameter.Formula);
                lines.Add(parameter.IsReporting ? "Отчётный размер: зависит от геометрии." : parameter.IsReadOnly ? "Только чтение." : "");
                if (parameter.Inputs.Count > 0) lines.Add("Зависит от: " + string.Join(", ", parameter.Inputs));
                if (parameter.Dependents.Count > 0) lines.Add("Влияет на: " + string.Join(", ", parameter.Dependents.Take(5))
                    + (parameter.Dependents.Count > 5 ? "…" : ""));
                lines.AddRange(parameter.Links.Take(4));
                return string.Join("\n", lines);
            }
            private HashSet<string> Reachable(IEnumerable<string> seeds, bool upstream)
            {
                var visited = new HashSet<string>(StringComparer.Ordinal); var pending = new Stack<string>(seeds);
                while (pending.Count > 0)
                {
                    string name = pending.Pop(); ParameterDependency parameter;
                    if (!visited.Add(name) || !Parameters.TryGetValue(name, out parameter)) continue;
                    foreach (string next in upstream ? parameter.Inputs : parameter.Dependents) pending.Push(next);
                }
                return visited;
            }
            internal string Report(IEnumerable<string> names)
            {
                var seeds = names.SelectMany(FamilyParameterNames.Aliases).Where(Parameters.ContainsKey).Distinct().ToList(); var relevant = Reachable(seeds, false);
                // Включаем и остальные входы зависимых формул: например, флаги видимости и включения соединителей.
                relevant.UnionWith(Reachable(relevant.ToList(), true));
                var lines = new List<string> { "Зависимости параметров конфигуратора (снимок исходного семейства):" };
                foreach (string name in relevant.OrderBy(n => n, StringComparer.Ordinal))
                {
                    ParameterDependency parameter;
                    if (!Parameters.TryGetValue(name, out parameter)) continue;
                    lines.Add(name + (string.IsNullOrWhiteSpace(parameter.Formula) ? " [без формулы]" : " = " + parameter.Formula)
                        + " [ReadOnly=" + parameter.IsReadOnly + "; Reporting=" + parameter.IsReporting
                        + "; Instance=" + parameter.IsInstance + "; исходное значение=" + parameter.SourceValue + "]");
                    lines.AddRange(parameter.Links.Distinct().Select(link => "  " + link));
                }
                lines.AddRange(ReadWarnings);
                return string.Join("\n", lines);
            }
        }

        internal sealed class DimensionValue : EditableValue
        {
            private string _text;
            private string _sourceText;
            private double _sourceMillimeters;
            internal bool AllowsSigned { get; private set; }
            public bool IsCalculated { get; private set; }
            public bool IsReadOnly { get { return IsCalculated; } }
            public bool IsCurrent { get; private set; } = true;
            public string Text
            {
                get { return IsCalculated && !IsCurrent ? "—" : _text; }
                set { if (IsReadOnly || _text == value) return; _text = value; NotifyValue(); }
            }
            private void NotifyValue() { Changed(nameof(Text)); Changed(nameof(IsValid)); Changed(nameof(Display)); }
            internal void SetCalculatedMode(bool calculated)
            {
                if (IsCalculated == calculated) return;
                IsCalculated = calculated; IsCurrent = !calculated;
                Changed(nameof(IsCalculated)); Changed(nameof(IsReadOnly)); Changed(nameof(IsCurrent)); NotifyValue();
            }
            internal void Invalidate()
            { if (IsCalculated) { IsCurrent = false; Changed(nameof(IsCurrent)); NotifyValue(); } }
            internal void ReadResult(double millimeters)
            {
                if (!IsCalculated) return;
                ReadSource(millimeters);
                Changed(nameof(IsCurrent)); NotifyValue();
            }
            private void ReadSource(double millimeters)
            {
                _sourceMillimeters = millimeters;
                _text = _sourceText = IsFinite(millimeters) ? millimeters.ToString("0.######", CultureInfo.CurrentCulture) : string.Empty;
                IsCurrent = IsFinite(millimeters);
            }
            internal static bool IsPositive(double value)
            { return IsFinite(value) && value > 0; }
            internal static bool IsFinite(double value) { return !double.IsNaN(value) && !double.IsInfinity(value); }
            internal bool TryValue(out double value)
            { return TryNumber(out value) && (AllowsSigned || value > 0); }
            internal bool TryNumber(out double value)
            {
                value = 0;
                if (IsCalculated && !IsCurrent) return false;
                // Округление в поле ввода не должно превращаться в изменение параметра Revit.
                if (_text == _sourceText) { value = _sourceMillimeters; return IsFinite(value); }
                return double.TryParse((_text ?? "").Trim().Replace(',', '.'),
                    NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out value)
                    && IsFinite(value);
            }
            public bool IsValid { get { double value; return IsCalculated || TryValue(out value); } }
            public string Display
            { get { double value; return TryNumber(out value) ? value.ToString("0.######", CultureInfo.CurrentCulture) : "—"; } }
            internal DimensionValue(double millimeters, bool calculated = false, bool allowsSigned = false)
            { ReadSource(millimeters); IsCalculated = calculated; AllowsSigned = allowsSigned; }
            internal double Millimeters(string name)
            {
                double value;
                if (!TryValue(out value)) throw new ArgumentException("«" + name + "»: введите число" + (AllowsSigned ? "" : " больше нуля") + " в миллиметрах.");
                return value;
            }
            internal DimensionValue Copy()
            {
                return new DimensionValue(_sourceMillimeters, IsCalculated, AllowsSigned)
                { _text = _text, _sourceText = _sourceText, IsCurrent = IsCurrent };
            }
        }

        internal sealed class BooleanValue : EditableValue
        {
            private bool? _value;
            public bool IsCalculated { get; private set; }
            public bool IsEnabled { get { return !IsCalculated; } }
            public bool IsValid { get { return IsCalculated || _value.HasValue; } }
            public bool? Value
            {
                get { return _value; }
                set { if (IsCalculated || _value == value) return; _value = value; NotifyValue(); }
            }
            public int SideIndex
            {
                get { return !_value.HasValue ? -1 : _value.Value ? 0 : 1; }
                set { if (value == 0 || value == 1) Value = value == 0; }
            }
            private void NotifyValue() { Changed(nameof(Value)); Changed(nameof(SideIndex)); Changed(nameof(IsValid)); }
            internal BooleanValue(bool? value, bool calculated = false) { _value = value; IsCalculated = calculated; }
            internal void SetCalculatedMode(bool calculated)
            { if (IsCalculated == calculated) return; IsCalculated = calculated; Changed(nameof(IsCalculated)); Changed(nameof(IsEnabled)); Changed(nameof(IsValid)); }
            internal void ReadResult(bool? value) { if (_value == value) return; _value = value; NotifyValue(); }
            internal BooleanValue Copy() { return new BooleanValue(_value, IsCalculated); }
        }

        internal sealed class SectionValue : EditableValue
        {
            private SectionTypeChoice _type;
            internal bool IsConnector, IsValve;
            internal DimensionValue Width, Height, OffsetX, OffsetY;
            internal DimensionValue Length;
            internal bool IsFixed { get { return IsConnector || IsValve; } }
            public SectionTypeChoice Type
            {
                get { return _type; }
                set
                {
                    if (ReferenceEquals(_type, value)) return;
                    bool changed = _type?.Key != value?.Key;
                    _type = value;
                    if (changed)
                    {
                        ResetEmptyLength();
                        Changed();
                    }
                }
            }
            internal void ResetEmptyLength()
            {
                // Начальное значение при выборе пустого блока. Дальше длина редактируется как обычно.
                if (Type != null && Type.IsEmpty && Length != null) Length.Text = "1";
            }
            internal SectionValue Copy()
            {
                return new SectionValue
                {
                    _type = _type,
                    IsConnector = IsConnector,
                    IsValve = IsValve,
                    Length = Length.Copy(),
                    Width = Width?.Copy(),
                    Height = Height?.Copy(),
                    OffsetX = OffsetX?.Copy(),
                    OffsetY = OffsetY?.Copy()
                };
            }
        }

        internal sealed class InstallationInfo : EditableValue
        {
            internal const string SystemNameParameter = "КП_О_Имя Системы";
            internal const string ManufacturerParameter = "КП_О_Завод-изготовитель";
            internal const string MarkParameter = "КП_О_Марка";
            internal const string UnitParameter = "КП_О_Единица измерения";
            internal const string DescriptionParameter = "КП_О_Наименование";
            internal const string ProductCodeParameter = "КП_О_Код изделия", MassTextParameter = "КП_О_Масса_Текст";
            private string _systemName = "", _manufacturer = "", _mark = "", _unit = "", _description = "", _productCode = "", _massText = "";
            private void Set(ref string field, string value, [CallerMemberName] string property = null)
            { value = value ?? ""; if (field == value) return; field = value; Changed(property); }
            public string SystemName { get { return _systemName; } set { Set(ref _systemName, value); } }
            public string Manufacturer { get { return _manufacturer; } set { Set(ref _manufacturer, value); } }
            public string Mark { get { return _mark; } set { Set(ref _mark, value); } }
            public string Unit { get { return _unit; } set { Set(ref _unit, value); } }
            public string Description { get { return _description; } set { Set(ref _description, value); } }
            public string ProductCode { get { return _productCode; } set { Set(ref _productCode, value); } }
            public string MassText { get { return _massText; } set { Set(ref _massText, value); } }
            internal string AutomaticTypeName
            {
                get
                {
                    return string.IsNullOrWhiteSpace(SystemName) || string.IsNullOrWhiteSpace(Mark)
                ? "Новый тип" : SystemName.Trim() + "_" + Mark.Trim();
                }
            }
            internal Dictionary<string, string> Assignments()
            {
                return new Dictionary<string, string> { { SystemNameParameter, SystemName }, { ManufacturerParameter, Manufacturer },
                { MarkParameter, Mark }, { UnitParameter, Unit }, { DescriptionParameter, Description },
                { ProductCodeParameter, ProductCode }, { MassTextParameter, MassText } };
            }
            internal InstallationInfo Copy()
            { return new InstallationInfo { SystemName = SystemName, Manufacturer = Manufacturer, Mark = Mark, Unit = Unit, Description = Description, ProductCode = ProductCode, MassText = MassText }; }
        }

        internal sealed class FamilyTypeItem : EditableValue
        {
            public bool IsCreate { get; internal set; }
            internal InstallationConfiguration Configuration;
            internal int SelectedSectionIndex, SelectedTab;
            internal string SavedPath, SourcePath, PersistedName, LoadError;
            internal bool HasAutomaticName;
            internal bool IsDirty = true;
            private string _savedSignature;
            private string StateSignature()
            { string name = (Name ?? "").Trim(); return name.Length + ":" + name + Configuration?.Signature(); }
            internal void MarkSaved() { _savedSignature = StateSignature(); IsDirty = false; }
            internal void MarkChanged() { IsDirty = _savedSignature == null || _savedSignature != StateSignature(); }
            internal void RequireSave() { _savedSignature = null; IsDirty = true; }
            internal void UpdateAutomaticName(IEnumerable<string> existingNames)
            { if (HasAutomaticName) Name = AvailableAutomaticName(Configuration.Info.AutomaticTypeName, existingNames); }
            private string _name;
            public string Name { get { return _name; } set { _name = value; Changed(); Changed(nameof(DisplayName)); } }
            public string DisplayName { get { return IsCreate ? "Создать тип" : string.IsNullOrWhiteSpace(Name) ? "Без имени" : Name; } }

            internal static string AvailableAutomaticName(string requested, IEnumerable<string> existingNames)
            {
                var names = new HashSet<string>(existingNames.Select(n => (n ?? "").Trim()), StringComparer.OrdinalIgnoreCase);
                if (!names.Contains(requested)) return requested;
                int number = 1;
                while (names.Contains(requested + " " + number)) number++;
                return requested + " " + number;
            }

            internal static FamilyTypeItem CreateDraft(IEnumerable<FamilyTypeItem> existing, SectionCatalog catalog)
            {
                return new FamilyTypeItem
                {
                    Name = AvailableAutomaticName("Новый тип", existing.Where(t => !t.IsCreate).Select(t => t.Name)),
                    HasAutomaticName = true,
                    Configuration = InstallationConfiguration.CreateDraft(catalog, newTypeDefaults: true)
                };
            }
        }

        // Правила из отчёта исходного RFA. Здесь нет Revit API, открытия документов и геометрического решателя.
        // Накопленные длины — результат сложения только видимых секций; длина последнего соединителя является входом.
        internal sealed class LocalGeometry
        {
            internal readonly Dictionary<int, double?> EndsMm = new Dictionary<int, double?>();
            internal string Limitation;
            internal double? TotalLengthMm { get { double? value; return EndsMm.TryGetValue(12, out value) ? value : null; } }
            internal static string Format(double? value)
            { return value.HasValue ? value.Value.ToString("0.######", CultureInfo.CurrentCulture) : "—"; }
            internal string Report()
            {
                return "Локальный расчёт по формулам исходного семейства:\n"
                    + string.Join("\n", EndsMm.OrderBy(p => p.Key).Select(p => "Системный_" + p.Key + "_Блок_Длина = " + Format(p.Value) + " мм"))
                    + "\nУстановка_Длина = " + Format(TotalLengthMm) + " мм"
                    + (string.IsNullOrWhiteSpace(Limitation) ? "" : "\n" + Limitation);
            }
            internal static bool SourceFlag(SectionCatalog catalog, string name, bool fallback)
            {
                ParameterDependency p;
                if (!catalog.Dependencies.Parameters.TryGetValue(name, out p)) return fallback;
                string formula = Compact(p.Formula);
                if (formula == "1=1") return true;
                if (formula == "1=0") return false;
                return p.IntegerValue.HasValue ? p.IntegerValue.Value != 0 : fallback;
            }
            private static string Compact(string formula) { return Regex.Replace(FamilyParameterNames.Canonical(formula), @"\s+", ""); }
            internal static LocalGeometry Calculate(InstallationConfiguration configuration)
            {
                var result = new LocalGeometry();
                var dependencies = configuration.Catalog.Dependencies;
                var issues = new List<string>();
                Action<string, string> check = (name, expected) =>
                {
                    ParameterDependency p;
                    if (dependencies.Parameters.TryGetValue(name, out p) && Compact(p.Formula) != Compact(expected))
                        issues.Add("Формула «" + name + "» отличается от перенесённой в плагин: " + (p.Formula ?? "без формулы") + ".");
                };
                check("Системный_1_Блок_Длина", "if(Соединитель_Приточный, Секция_1_Соединитель_Длина, 0 мм)");
                check("Системный_2_Блок_Длина", "if(Соединитель_Приточный_Клапан, Секция_2_Соединитель_Клапан_Длина + Системный_1_Блок_Длина, Системный_1_Блок_Длина)");
                for (int slot = 3; slot <= 11; slot++)
                {
                    int section = slot - 2;
                    check("Системный_" + slot + "_Блок_Длина", "if(Системный_Видимость_" + section + "_Блок, "
                        + InstallationConfiguration.DimensionParameterName(slot, "Длина") + " + Системный_" + (slot - 1)
                        + "_Блок_Длина, Системный_" + (slot - 1) + "_Блок_Длина)");
                    if (section > 1) check("Системный_Видимость_" + section + "_Блок",
                        "if(Секции_Промежуточные_Количество > " + (section - 1) + ", 1 = 1, 1 = 0)");
                }
                check("Системный_12_Блок_Длина", "if(Соединитель_Отработанный, Секция_12_Соединитель_Длина + Системный_11_Блок_Длина, Системный_11_Блок_Длина)");
                check("Установка_Длина", "Системный_12_Блок_Длина");
                foreach (string name in new[] { "Соединитель_Приточный", "Соединитель_Отработанный", "Системный_Видимость_1_Блок", "Соединитель_Приточный_Клапан" })
                {
                    ParameterDependency p;
                    if (dependencies.Parameters.TryGetValue(name, out p) && !string.IsNullOrWhiteSpace(p.Formula)
                        && Compact(p.Formula) != "1=1" && Compact(p.Formula) != "1=0")
                        issues.Add("Не перенесена формула «" + name + "»: " + p.Formula + ".");
                }
                var blocks = Enumerable.Range(0, configuration.Blocks.Count).ToDictionary(configuration.SlotAt, i => configuration.Blocks[i]);
                double? end = issues.Count == 0 ? (double?)0 : null;
                for (int slot = 1; slot <= 12; slot++)
                {
                    bool visible = slot == 1 ? SourceFlag(configuration.Catalog, "Соединитель_Приточный", true)
                        : slot == 12 ? SourceFlag(configuration.Catalog, "Соединитель_Отработанный", true)
                        : slot == 2 ? configuration.HasValve : slot == 3 ? SourceFlag(configuration.Catalog, "Системный_Видимость_1_Блок", true)
                        : slot - 2 <= configuration.IntermediateCount;
                    if (visible)
                    {
                        SectionValue block; double length;
                        if (!blocks.TryGetValue(slot, out block) || !block.Length.TryNumber(out length) || length < 0) end = null;
                        else if (end.HasValue) end += length;
                    }
                    result.EndsMm[slot] = end;
                }
                var unknown = configuration.Dimensions().Where(p => p.Value.IsCalculated).Select(p => p.Key).ToList();
                if (unknown.Count > 0) issues.Add("Для этих размеров локальные правила не подтверждены: " + string.Join(", ", unknown) + ".");
                result.Limitation = string.Join("\n", issues);
                return result;
            }
        }

        internal static class ValveControl
        {
            internal const string ParameterName = "Соединитель_Приточный_Клапан";
            internal static bool? ConstantFormula(string formula)
            {
                string value = Regex.Replace(formula ?? "", @"\s+", "");
                return value == "1=1" ? (bool?)true : value == "1=0" ? (bool?)false : null;
            }
            internal static bool CanEdit(ParameterDependency parameter)
            {
                return parameter == null || !parameter.IsReadOnly && !parameter.IsReporting
                && (string.IsNullOrWhiteSpace(parameter.Formula) || ConstantFormula(parameter.Formula).HasValue);
            }
        }

        internal sealed class InstallationConfiguration
        {
            private string _calculationInputSignature;
            internal List<SectionValue> Blocks = new List<SectionValue>();
            // Каталог содержит имена и исходные размеры, никаких ElementId из закрытого документа.
            internal SectionCatalog Catalog;
            internal InstallationInfo Info = new InstallationInfo();
            // Один набор значений на тип установки: все подходящие секции используют те же объекты.
            internal Dictionary<string, DimensionValue> SharedDimensions = new Dictionary<string, DimensionValue>();
            internal Dictionary<string, BooleanValue> SharedBooleans = new Dictionary<string, BooleanValue>();
            internal DimensionValue InstallationWidth, InstallationHeight, FrameHeight;
            internal LocalGeometry Geometry = new LocalGeometry();
            internal bool CanChangeValve { get { return ValveControl.CanEdit(Catalog.Dependencies.Find(ValveControl.ParameterName)); } }
            private SectionValue _valve;
            internal bool HasValve { get { return Blocks.Count > 1 && Blocks[1].IsValve; } }
            internal int Minimum { get { return HasValve ? 4 : 3; } }
            internal int Capacity { get { return HasValve ? 12 : 11; } }
            internal int SectionCount { get { return Blocks.Count; } }
            internal int IntermediateCount { get { return Blocks.Count(b => !b.IsFixed); } }
            internal int FirstSectionIndex { get { return HasValve ? 2 : 1; } }
            private static int SlotAt(int index, IList<SectionValue> blocks)
            {
                if (index == 0) return 1;
                if (index == blocks.Count - 1) return 12;
                if (blocks[index].IsValve) return 2;
                return index + (blocks[1].IsValve ? 1 : 2);
            }
            internal int SlotAt(int index) { return SlotAt(index, Blocks); }
            internal string BlockTitle(int index)
            {
                if (Blocks[index].IsConnector) return "Соединитель";
                if (Blocks[index].IsValve) return "Клапан";
                return "Секция " + (index - FirstSectionIndex + 1);
            }
            internal static string ParameterName(int slot)
            {
                if (slot < 1 || slot > 12) throw new ArgumentOutOfRangeException(nameof(slot));
                if (slot == 1 || slot == 12) return "Секция_" + slot + "_Соединитель_Тип";
                if (slot == 2) return "Секция_2_Соединитель_Клапан_Тип";
                return "Секция_" + slot + "_Промежуточная_" + (slot - 2) + "_Тип";
            }
            internal static string DimensionParameterName(int slot, string dimension)
            {
                if (dimension != "Длина" && dimension != "Ширина" && dimension != "Высота" && dimension != "Смещение по X" && dimension != "Смещение по Y")
                    throw new ArgumentException("Неизвестный размер секции.");
                if (dimension != "Длина" && slot != 1 && slot != 12)
                    throw new ArgumentException("Ширина, высота и смещения задаются только для соединителей.");
                string typeParameter = ParameterName(slot);
                return typeParameter.Substring(0, typeParameter.Length - "Тип".Length) + dimension;
            }
            internal static IEnumerable<string> InterfaceParameterNames()
            {
                foreach (string name in new InstallationInfo().Assignments().Keys) yield return name;
                foreach (string name in SharedParameters.LengthNames.Concat(SharedParameters.BooleanNames)) yield return name;
                yield return "Секции_Промежуточные_Количество"; yield return "Соединитель_Приточный_Клапан";
                yield return "Установка_Ширина"; yield return "Установка_Высота"; yield return "Основание_Рама_Высота";
                for (int slot = 1; slot <= 12; slot++)
                {
                    yield return ParameterName(slot); yield return DimensionParameterName(slot, "Длина");
                    if (slot != 1 && slot != 12) continue;
                    yield return DimensionParameterName(slot, "Ширина"); yield return DimensionParameterName(slot, "Высота");
                    yield return DimensionParameterName(slot, "Смещение по X"); yield return DimensionParameterName(slot, "Смещение по Y");
                }
            }
            internal static InstallationConfiguration CreateDraft(SectionCatalog catalog, bool newTypeDefaults = false)
            {
                if (catalog == null || catalog.Count != 12 || !Enumerable.Range(1, 12).All(catalog.ContainsKey))
                    throw new InvalidOperationException("Не загружен состав секций исходного семейства.");
                if (catalog.IntermediateCount < 1 || catalog.IntermediateCount > 9)
                    throw new InvalidOperationException("В исходном типе некорректное количество промежуточных секций.");
                var configuration = new InstallationConfiguration
                {
                    Catalog = catalog,
                    InstallationWidth = catalog.Dimension("Установка_Ширина", catalog.InstallationWidthMm),
                    InstallationHeight = catalog.Dimension("Установка_Высота", catalog.InstallationHeightMm),
                    FrameHeight = catalog.Dimension("Основание_Рама_Высота", catalog.FrameHeightMm)
                };
                foreach (string name in SharedParameters.LengthNames)
                {
                    double value;
                    configuration.SharedDimensions.Add(name, catalog.Dimension(name, catalog.SharedLengthsMm.TryGetValue(name, out value) ? value : double.NaN));
                }
                foreach (string name in SharedParameters.BooleanNames)
                {
                    bool? value;
                    configuration.SharedBooleans.Add(name, new BooleanValue(catalog.SharedFlags.TryGetValue(name, out value) ? value : null,
                        catalog.Dependencies.IsCalculated(name)));
                }
                configuration._valve = configuration.NewBlock(2, valve: true, resetEmptyLength: false);
                configuration.Blocks.Add(configuration.NewBlock(1, connector: true, resetEmptyLength: false));
                if (catalog.HasValve) configuration.Blocks.Add(configuration._valve);
                for (int slot = 3; slot < 3 + catalog.IntermediateCount; slot++)
                    configuration.Blocks.Add(configuration.NewBlock(slot, resetEmptyLength: false));
                configuration.Blocks.Add(configuration.NewBlock(12, connector: true, resetEmptyLength: false));
                configuration.RecalculateLocal();
                if (newTypeDefaults)
                {
                    configuration.SharedDimensions[SharedParameters.ServiceDepth].Text = "1000";
                    configuration.SharedBooleans[SharedParameters.ServiceRight].Value = true;
                    foreach (var block in configuration.Blocks.Concat(new[] { configuration._valve }).Distinct()) block.ResetEmptyLength();
                    configuration.RecalculateLocal();
                }
                return configuration;
            }
            private SectionValue NewBlock(int slot, bool connector = false, bool valve = false, bool resetEmptyLength = true)
            {
                var definition = Catalog[slot];
                var block = new SectionValue
                {
                    IsConnector = connector,
                    IsValve = valve,
                    Type = definition.SourceType,
                    Length = Catalog.Dimension(DimensionParameterName(slot, "Длина"), definition.DefaultLengthMm),
                    Width = connector ? Catalog.Dimension(DimensionParameterName(slot, "Ширина"), definition.DefaultWidthMm) : null,
                    Height = connector ? Catalog.Dimension(DimensionParameterName(slot, "Высота"), definition.DefaultHeightMm) : null,
                    OffsetX = connector ? Catalog.Dimension(DimensionParameterName(slot, "Смещение по X"), definition.DefaultOffsetXMm) : null,
                    OffsetY = connector ? Catalog.Dimension(DimensionParameterName(slot, "Смещение по Y"), definition.DefaultOffsetYMm) : null
                };
                if (resetEmptyLength) block.ResetEmptyLength();
                return block;
            }
            internal InstallationConfiguration Copy()
            {
                var copy = new InstallationConfiguration
                {
                    Catalog = Catalog,
                    Info = Info.Copy(),
                    SharedDimensions = SharedDimensions.ToDictionary(p => p.Key, p => p.Value.Copy()),
                    SharedBooleans = SharedBooleans.ToDictionary(p => p.Key, p => p.Value.Copy()),
                    InstallationWidth = InstallationWidth.Copy(),
                    InstallationHeight = InstallationHeight.Copy(),
                    FrameHeight = FrameHeight.Copy(),
                    Blocks = Blocks.Select(b => b.Copy()).ToList(),
                    _calculationInputSignature = _calculationInputSignature
                };
                copy._valve = HasValve ? copy.Blocks[1] : _valve.Copy();
                copy.RecalculateLocal();
                return copy;
            }
            internal void SetValve(bool enabled)
            {
                if (enabled == HasValve) return;
                if (!CanChangeValve) throw new InvalidOperationException("Наличие клапана задано формулой или свойствами исходного семейства.\n"
                    + Catalog.Dependencies.Tooltip("Соединитель_Приточный_Клапан"));
                var candidate = Blocks.ToList();
                if (enabled) candidate.Insert(1, _valve);
                else candidate.RemoveAt(1);
                ValidateChoices(candidate); Blocks = candidate; RefreshDimensionModes();
                // При выключении сохраняем выбранный состав клапана для следующего включения.
            }
            internal void Resize(int count)
            {
                if (count < Minimum || count > Capacity)
                    throw new ArgumentException("Количество блоков должно быть от " + Minimum + " до " + Capacity + ".");
                var candidate = Blocks.ToList();
                while (candidate.Count > count) candidate.RemoveAt(candidate.Count - 2);
                while (candidate.Count < count)
                    candidate.Insert(candidate.Count - 1, NewBlock(3 + candidate.Count(b => !b.IsFixed)));
                ValidateChoices(candidate); Blocks = candidate; RefreshDimensionModes();
            }
            internal void InsertAfter(int index)
            {
                if (SectionCount >= Capacity) throw new InvalidOperationException("Достигнут предел секций исходного семейства.");
                int target = Math.Max(FirstSectionIndex, Math.Min(index + 1, Blocks.Count - 1));
                var candidate = Blocks.ToList(); candidate.Insert(target, NewBlock(target + (HasValve ? 1 : 2)));
                ValidateChoices(candidate); Blocks = candidate; RefreshDimensionModes();
            }
            internal void Remove(int index)
            {
                if (index < FirstSectionIndex || index >= Blocks.Count - 1 || SectionCount <= Minimum) return;
                var candidate = Blocks.ToList(); candidate.RemoveAt(index);
                ValidateChoices(candidate); Blocks = candidate; RefreshDimensionModes();
            }
            internal void Move(int index, int target)
            {
                if (index < FirstSectionIndex || target < FirstSectionIndex || index >= Blocks.Count - 1 || target >= Blocks.Count - 1) return;
                var candidate = Blocks.ToList(); var block = candidate[index]; candidate.RemoveAt(index); candidate.Insert(target, block);
                ValidateChoices(candidate); Blocks = candidate; RefreshDimensionModes();
            }
            private void ValidateChoices(IList<SectionValue> blocks)
            {
                for (int i = 0; i < blocks.Count; i++)
                {
                    int slot = SlotAt(i, blocks);
                    if (blocks[i].Type == null || !Catalog[slot].Choices.Any(c => c.Key == blocks[i].Type.Key))
                        throw new InvalidOperationException("Выбранный состав недоступен для параметра «" + Catalog[slot].ParameterName + "».");
                    if (blocks[i].IsConnector && !blocks[i].Type.IsConnectorType)
                        throw new InvalidOperationException("Для соединителей доступны гибкая вставка, воздушный клапан, шумоглушитель и пустой блок.");
                }
            }
            internal void Validate()
            {
                if (SectionCount < Minimum || SectionCount > Capacity || IntermediateCount < 1 || IntermediateCount > 9
                    || !Blocks[0].IsConnector || !Blocks.Last().IsConnector
                    || Blocks.Count(b => b.IsConnector) != 2 || Blocks.Count(b => b.IsValve) != (HasValve ? 1 : 0))
                    throw new ArgumentException("Проверьте количество секций и положение клапана.");
                ValidateChoices(Blocks);
                ValidateDimensions();
                foreach (var flag in SharedBooleans)
                    if (!flag.Value.IsValid) throw new ArgumentException("Задайте значение «" + SharedParameters.Label(flag.Key) + "».");
                if (!CanChangeValve && HasValve != LocalGeometry.SourceFlag(Catalog, "Соединитель_Приточный_Клапан", true))
                    throw new InvalidOperationException("Состояние клапана не соответствует формуле исходного семейства.");
            }
            private void ValidateDimensions()
            {
                foreach (var dimension in Dimensions())
                    if (!dimension.Value.IsCalculated) dimension.Value.Millimeters(dimension.Key);
            }
            internal string Signature()
            {
                // Сравниваем сохраняемые входы, а не результаты формул или ссылки на объекты ComboBox.
                var values = Dimensions().Where(p => !p.Value.IsCalculated).OrderBy(p => p.Key, StringComparer.Ordinal).Select(p =>
                { double value; return p.Key + "=" + (p.Value.TryNumber(out value) ? value.ToString("R", CultureInfo.InvariantCulture) : p.Value.Text); });
                var fields = new[] { HasValve.ToString() }.Concat(Blocks.Select(b => b.Type?.Key ?? ""))
                    .Concat(values).Concat(SharedBooleans.Where(p => !p.Value.IsCalculated).OrderBy(p => p.Key, StringComparer.Ordinal)
                        .Select(p => p.Key + "=" + p.Value.Value))
                    .Concat(Info.Assignments().OrderBy(p => p.Key, StringComparer.Ordinal).Select(p => p.Key + "=" + p.Value));
                return string.Concat(fields.Select(value => value.Length + ":" + value));
            }
            internal void ValidateForOutput()
            {
                var missing = IncompleteFields();
                if (missing.Count > 0) throw new InvalidOperationException("Заполните параметры установки:\n" + string.Join("\n", missing));
                Validate();
            }
            internal List<string> IncompleteFields()
            {
                var fields = new List<string>();
                Action<string, DimensionValue> check = (label, value) => { if (value == null || !value.IsValid) fields.Add(label); };
                foreach (var dimension in SharedDimensions) check(SharedParameters.Label(dimension.Key), dimension.Value);
                foreach (var flag in SharedBooleans) if (!flag.Value.IsValid) fields.Add(SharedParameters.Label(flag.Key));
                check("Ширина установки", InstallationWidth); check("Высота установки", InstallationHeight); check("Высота рамы", FrameHeight);
                for (int i = 0; i < Blocks.Count; i++)
                {
                    var block = Blocks[i];
                    string title = block.IsConnector ? (i == 0 ? "Первый соединитель" : "Последний соединитель") : BlockTitle(i);
                    if (block.Type == null) fields.Add(title + ": тип оборудования");
                    else if (block.IsConnector && !block.Type.IsConnectorType) fields.Add(title + ": выберите гибкую вставку, воздушный клапан, шумоглушитель или пустой блок");
                    check(title + ": длина", block.Length);
                    if (!block.IsConnector) continue;
                    check(title + ": ширина", block.Width); check(title + ": высота", block.Height);
                    check(title + ": смещение по X", block.OffsetX); check(title + ": смещение по Y", block.OffsetY);
                }
                return fields;
            }
            internal Dictionary<string, DimensionValue> Dimensions()
            {
                var result = new Dictionary<string, DimensionValue> { { "Установка_Ширина", InstallationWidth },
                    { "Установка_Высота", InstallationHeight }, { "Основание_Рама_Высота", FrameHeight } };
                foreach (var dimension in SharedDimensions) result.Add(dimension.Key, dimension.Value);
                for (int i = 0; i < Blocks.Count; i++)
                {
                    var block = Blocks[i]; int slot = SlotAt(i);
                    result[DimensionParameterName(slot, "Длина")] = block.Length;
                    if (!block.IsConnector) continue;
                    result[DimensionParameterName(slot, "Ширина")] = block.Width;
                    result[DimensionParameterName(slot, "Высота")] = block.Height;
                    result[DimensionParameterName(slot, "Смещение по X")] = block.OffsetX;
                    result[DimensionParameterName(slot, "Смещение по Y")] = block.OffsetY;
                }
                return result;
            }
            internal void RefreshDimensionModes(Func<string, bool> isCalculated = null)
            {
                foreach (var dimension in Dimensions())
                    dimension.Value.SetCalculatedMode(isCalculated == null ? Catalog.Dependencies.IsCalculated(dimension.Key) : isCalculated(dimension.Key));
            }
            internal void RefreshBooleanModes(Func<string, bool> isCalculated = null)
            {
                foreach (var flag in SharedBooleans)
                    flag.Value.SetCalculatedMode(isCalculated == null ? Catalog.Dependencies.IsCalculated(flag.Key) : isCalculated(flag.Key));
            }
            internal Dictionary<string, int> BooleanAssignments()
            {
                return SharedBooleans.Where(p => !p.Value.IsCalculated && p.Value.Value.HasValue)
                .ToDictionary(p => p.Key, p => p.Value.Value.Value ? 1 : 0);
            }
            internal void RecalculateLocal()
            {
                RefreshDimensionModes(); RefreshBooleanModes();
                string inputs = HasValve + "|" + string.Join("|", Blocks.Select(b => b.Type?.Key ?? ""))
                    + "|" + string.Join("|", Dimensions().Where(p => !p.Value.IsCalculated).OrderBy(p => p.Key, StringComparer.Ordinal)
                        .Select(p => p.Key + "=" + p.Value.Text))
                    + "|" + string.Join("|", SharedBooleans.Where(p => !p.Value.IsCalculated).OrderBy(p => p.Key, StringComparer.Ordinal)
                        .Select(p => p.Key + "=" + p.Value.Value));
                // При чтении показываем результаты исходного типа. Они устаревают только после изменения входов.
                if (_calculationInputSignature != null && _calculationInputSignature != inputs)
                    foreach (var dimension in Dimensions().Values) dimension.Invalidate();
                _calculationInputSignature = inputs;
                Geometry = LocalGeometry.Calculate(this);
            }
            internal DimensionValue BlockWidth(int index)
            { return Blocks[index].IsConnector ? Blocks[index].Width : Blocks[index].IsValve ? Blocks[0].Width : InstallationWidth; }
            internal DimensionValue BlockHeight(int index)
            { return Blocks[index].IsConnector ? Blocks[index].Height : Blocks[index].IsValve ? Blocks[0].Height : InstallationHeight; }
            private IEnumerable<int> InactiveLengthSlots()
            {
                var active = new HashSet<int>(Enumerable.Range(0, Blocks.Count).Select(SlotAt));
                for (int slot = 1; slot <= 12; slot++)
                    if (!active.Contains(slot) && !Catalog.Dependencies.IsCalculated(DimensionParameterName(slot, "Длина")))
                        yield return slot;
            }
            internal bool NeedsInactiveLengthReset
            {
                get
                {
                    return InactiveLengthSlots().Any(slot => !DimensionValue.IsFinite(Catalog[slot].DefaultLengthMm)
                    || Math.Abs(Catalog[slot].DefaultLengthMm - 1.0) > 0.000001);
                }
            }
            internal Dictionary<string, double> DimensionAssignments()
            {
                Validate();
                var result = new Dictionary<string, double>();
                foreach (var dimension in Dimensions())
                    if (!dimension.Value.IsCalculated) result[dimension.Key] = dimension.Value.Millimeters(dimension.Key);
                    else result.Remove(dimension.Key);
                // Отключённый клапан и скрытые промежуточные позиции имеют длину 1 мм.
                // Видимые блоки сохраняют заданную длину, в том числе вручную изменённый пустой блок.
                foreach (int slot in InactiveLengthSlots()) result[DimensionParameterName(slot, "Длина")] = 1.0;
                return result;
            }
            internal Dictionary<int, SectionTypeChoice> SectionAssignments()
            {
                Validate();
                var result = new Dictionary<int, SectionTypeChoice>();
                for (int index = 0; index < Blocks.Count; index++) result[SlotAt(index)] = Blocks[index].Type;
                return result;
            }
        }

        internal sealed class FamilyActionPlan
        {
            internal string Path;
            internal bool WritesFamily, NeedsPath;
            internal static FamilyActionPlan Make(RequestKind kind, string savedPath, string configuredPath, string persistedName, bool dirty)
            {
                bool write = kind == RequestKind.Save || kind == RequestKind.DeleteType;
                if (!write && kind != RequestKind.OpenInRevit && kind != RequestKind.AddToProject) throw new ArgumentException("Неизвестное действие.");
                if (!write && (dirty || string.IsNullOrWhiteSpace(savedPath) || string.IsNullOrWhiteSpace(persistedName)))
                    throw new InvalidOperationException("Сначала сохраните тип и изменения кнопкой «Сохранить».");
                // Для определённого проекта место записи всегда задаёт его папка.
                // Открытие уже сохранённого типа использует его фактический файл.
                string path = write && !string.IsNullOrWhiteSpace(configuredPath) ? configuredPath : savedPath;
                return new FamilyActionPlan { WritesFamily = write, Path = path, NeedsPath = write && string.IsNullOrWhiteSpace(path) };
            }
        }

        internal sealed class FamilyPackage
        {
            internal SectionCatalog Catalog;
            internal List<FamilyTypeItem> Types = new List<FamilyTypeItem>();
            internal string Path, SourcePath, SelectedTypeName;
        }

        // Создание и редактирование различаются намерением, а не совпадением введённого имени.
        internal sealed class TypeWritePlan
        {
            internal string SourceTypeName, TargetName;
            internal bool Create;
            internal static TypeWritePlan Make(IEnumerable<string> existingNames, string targetName, string persistedName, bool freshLibraryCopy)
            {
                var names = existingNames.ToList();
                var result = new TypeWritePlan { TargetName = targetName, Create = string.IsNullOrWhiteSpace(persistedName) };
                if (string.Equals(targetName, BaseFamilyTypeName, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Шаблонный тип нельзя перезаписывать. Задайте имя установки.");
                if (string.Equals(persistedName, BaseFamilyTypeName, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Шаблонный тип служит только основой для новой установки.");
                if (!result.Create)
                {
                    result.SourceTypeName = names.SingleOrDefault(n => string.Equals(n, persistedName, StringComparison.Ordinal));
                    if (result.SourceTypeName == null) throw new InvalidOperationException("Редактируемый тип «" + persistedName + "» не найден. Заново откройте семейство.");
                }
                else
                {
                    result.SourceTypeName = names.SingleOrDefault(n => string.Equals(n, BaseFamilyTypeName, StringComparison.Ordinal));
                    if (result.SourceTypeName == null) throw new InvalidOperationException("В семействе нет шаблонного типа «" + BaseFamilyTypeName + "» для создания новой установки.");
                }
                if (names.Any(n => (result.Create || n != result.SourceTypeName)
                    && string.Equals(n, targetName, StringComparison.OrdinalIgnoreCase)))
                    throw new InvalidOperationException("Тип «" + targetName + "» уже существует. Выберите его в списке для редактирования или задайте другое имя новому типу.");
                return result;
            }
        }
        internal sealed class OperationError
        {
            internal string Summary, Details;
            internal static string Message(Exception error)
            {
                for (Exception current = error; current != null; current = current.InnerException)
                    if (!string.IsNullOrWhiteSpace(current.Message))
                        return string.Join(" ", current.Message.Split((char[])null, StringSplitOptions.RemoveEmptyEntries));
                return error.GetType().FullName;
            }
            internal static OperationError Create(string stage, Exception error, string savedPath, bool writeAttempt = true)
            {
                string saved = savedPath == null ? "" : "Копия сохранена. ";
                return new OperationError
                {
                    Summary = saved + stage + ": " + Message(error),
                    Details = "Этап: " + stage + "\n" + (savedPath == null ? (writeAttempt ? "Копия не сохранена.\n" : "Сохранение файла не выполнялось.\n")
                        : "Копия сохранена: " + savedPath + "\n") + "\n" + error.ToString()
                };
            }
        }
        // CONFIGURATION MODEL END

        internal sealed class OpenDocumentItem
        {
            internal Document Document { get; set; }
            public string DisplayName { get; set; }
            public string LoadTargetName { get; set; }
        }

        internal sealed class ProjectItem
        {
            internal string Name { get; set; }
            internal string Stage { get; set; }
            internal string Code { get; set; }
            internal string[] Paths { get; set; }
            internal bool IsUnknown { get; set; }
            public string DisplayName
            {
                get { return IsUnknown ? UnknownProjectName : Name + " [" + Stage + "]"; }
            }
            public string Description
            {
                get
                {
                    return DisplayName + (Paths == null || Paths.Length == 0
                    ? string.Empty : "\n" + string.Join("\n", Paths));
                }
            }
        }

        internal sealed class ProjectMatch
        {
            internal ProjectItem Project { get; set; }
            internal string RootPath { get; set; }
            internal bool IsAmbiguous { get; set; }
        }

        internal static class ProjectFamilyStorage
        {
            internal static string GetPath(ProjectItem project)
            {
                if (project == null || project.IsUnknown) return null;
                string code = CleanDatabaseValue(project.Code);
                if (code.Length == 0)
                    throw new InvalidOperationException("В базе проектов не задана аббревиатура проекта «" + project.DisplayName + "».");
                // Аббревиатура — ровно одна папка Windows. Не подменяем её другим именем.
                if (code == "." || code == ".." || code.EndsWith(".", StringComparison.Ordinal)
                    || code.Any(c => char.IsControl(c) || "<>:\"/\\|?*".IndexOf(c) >= 0)
                    || Regex.IsMatch(code, @"^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)", RegexOptions.IgnoreCase))
                    throw new InvalidOperationException("Аббревиатура проекта «" + code + "» содержит недопустимое имя папки.");
                return Path.Combine(ProjectFamiliesRoot, code, ProjectFamilyFileName);
            }

            internal static ProjectMatch Find(IList<ProjectItem> projects, string familyPath)
            {
                string normalized = ProjectPathMatcher.Normalize(familyPath);
                var result = new ProjectMatch();
                if (normalized.Length == 0) return result;
                foreach (var project in projects)
                {
                    string path;
                    try { path = GetPath(project); }
                    catch (InvalidOperationException) { continue; }
                    if (path == null || !string.Equals(normalized, ProjectPathMatcher.Normalize(path), StringComparison.OrdinalIgnoreCase)) continue;
                    if (result.Project != null && !ReferenceEquals(result.Project, project))
                        return new ProjectMatch { IsAmbiguous = true };
                    result.Project = project; result.RootPath = Path.GetDirectoryName(path);
                }
                return result;
            }

            internal static bool EnsureFile(string sourcePath, string path, Func<string, string, bool> confirm)
            {
                if (File.Exists(path)) return true;
                if (!confirm(sourcePath, path)) return false;
                // Пока был открыт вопрос, файл мог появиться у другого пользователя.
                if (File.Exists(path)) return true;
                if (!File.Exists(sourcePath))
                    throw new FileNotFoundException("Исходное семейство для копирования недоступно. Проверьте подключение к сети.", sourcePath);
                string directory = Path.GetDirectoryName(path);
                Directory.CreateDirectory(directory);
                string temporary = Path.Combine(directory, ".KPLN_Copy_" + Guid.NewGuid().ToString("N") + ".tmp");
                try
                {
                    // Готовый файл появляется только после полного копирования. Чужой файл не перезаписываем.
                    File.Copy(sourcePath, temporary, false);
                    File.SetAttributes(temporary, File.GetAttributes(temporary) & ~FileAttributes.ReadOnly);
                    try { File.Move(temporary, path); }
                    catch (IOException) { if (!File.Exists(path)) throw; }
                    return true;
                }
                finally
                {
                    try { if (File.Exists(temporary)) File.Delete(temporary); }
                    catch (IOException) { }
                    catch (UnauthorizedAccessException) { }
                }
            }
        }

        internal static string CleanDatabaseValue(string value)
        {
            string result = (value ?? string.Empty).Trim();
            return string.Equals(result, "Empty", StringComparison.OrdinalIgnoreCase) ? string.Empty : result;
        }

        // Чтение через SQLite из состава Windows 10/11: дополнительные DLL плагину не нужны.
        // https://www.sqlite.org/c3ref/open.html — SQLITE_OPEN_READONLY не создаёт базу.
        internal static class ProjectDatabase
        {
            private const string Query = "SELECT [Name], [Stage], [Code], [MainPath], " +
                "[RevitServerPath], [RevitServerPath2], [RevitServerPath3], [RevitServerPath4] FROM [Projects];";

            internal static IList<ProjectItem> Read(string path)
            {
                if (!File.Exists(path))
                    throw new FileNotFoundException("База проектов недоступна. Проверьте диск Z: и права на чтение.\n" + path);

                IntPtr database = IntPtr.Zero;
                IntPtr statement = IntPtr.Zero;
                try
                {
                    int result = sqlite3_open_v2(Encoding.UTF8.GetBytes(path + "\0"), out database, 1, IntPtr.Zero);
                    Check(result, database);
                    Check(sqlite3_busy_timeout(database, 1500), database);
                    byte[] sql = Encoding.UTF8.GetBytes(Query + "\0");
                    IntPtr tail;
                    Check(sqlite3_prepare_v2(database, sql, sql.Length, out statement, out tail), database);
                    var projects = new List<ProjectItem>();
                    while ((result = sqlite3_step(statement)) == 100) // SQLITE_ROW
                    {
                        var paths = new List<string>();
                        for (int column = 3; column < 8; column++)
                        {
                            string value = ReadText(statement, column);
                            if (value.Length != 0) paths.Add(value);
                        }
                        string name = ReadText(statement, 0);
                        string stage = ReadText(statement, 1);
                        projects.Add(new ProjectItem
                        {
                            Name = name.Length == 0 ? UnknownProjectName : name,
                            Stage = stage.Length == 0 ? UnknownProjectName : stage,
                            Code = ReadText(statement, 2),
                            Paths = paths.Distinct(StringComparer.OrdinalIgnoreCase).ToArray()
                        });
                    }
                    if (result != 101) Check(result, database); // SQLITE_DONE
                    return projects.OrderBy(p => p.DisplayName, StringComparer.CurrentCultureIgnoreCase).ToList();
                }
                catch (DllNotFoundException ex)
                {
                    throw new InvalidOperationException("Компонент SQLite Windows (winsqlite3.dll) недоступен.", ex);
                }
                finally
                {
                    if (statement != IntPtr.Zero) sqlite3_finalize(statement);
                    if (database != IntPtr.Zero) sqlite3_close(database);
                }
            }

            private static string ReadText(IntPtr statement, int column)
            {
                IntPtr pointer = sqlite3_column_text(statement, column);
                int length = sqlite3_column_bytes(statement, column);
                if (pointer == IntPtr.Zero || length == 0) return string.Empty;
                var bytes = new byte[length];
                Marshal.Copy(pointer, bytes, 0, length);
                return CleanDatabaseValue(Encoding.UTF8.GetString(bytes));
            }

            private static void Check(int result, IntPtr database)
            {
                if (result == 0) return;
                IntPtr pointer = database == IntPtr.Zero ? IntPtr.Zero : sqlite3_errmsg(database);
                var bytes = new List<byte>();
                if (pointer != IntPtr.Zero)
                    for (int i = 0; Marshal.ReadByte(pointer, i) != 0; i++)
                        bytes.Add(Marshal.ReadByte(pointer, i));
                throw new InvalidOperationException("Не удалось прочитать таблицу Projects. SQLite " + result + ": " +
                    Encoding.UTF8.GetString(bytes.ToArray()));
            }

            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
            private static extern int sqlite3_open_v2(byte[] filename, out IntPtr database, int flags, IntPtr vfs);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
            private static extern int sqlite3_busy_timeout(IntPtr database, int milliseconds);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
            private static extern int sqlite3_prepare_v2(IntPtr database, byte[] sql, int length, out IntPtr statement, out IntPtr tail);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
            private static extern int sqlite3_step(IntPtr statement);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
            private static extern IntPtr sqlite3_column_text(IntPtr statement, int column);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
            private static extern int sqlite3_column_bytes(IntPtr statement, int column);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
            private static extern IntPtr sqlite3_errmsg(IntPtr database);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
            private static extern int sqlite3_finalize(IntPtr statement);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
            private static extern int sqlite3_close(IntPtr database);
        }

        // Сопоставление только полных путей и границ папок: «Проект» не совпадает с «Проект2».
        internal static class ProjectPathMatcher
        {
            internal static string Normalize(string value)
            {
                string path = CleanDatabaseValue(value).Replace('\\', '/');
                string prefix;
                int protectedSegments;
                if (path.StartsWith("RSN://", StringComparison.OrdinalIgnoreCase))
                {
                    prefix = "RSN://";
                    path = path.Substring(6);
                    protectedSegments = 1; // имя Revit Server
                }
                else if (path.StartsWith("//", StringComparison.Ordinal))
                {
                    prefix = "//";
                    path = path.Substring(2);
                    protectedSegments = 2; // сервер и общая папка UNC
                }
                else if (path.Length >= 3 && char.IsLetter(path[0]) && path[1] == ':' && path[2] == '/')
                {
                    prefix = path.Substring(0, 3);
                    path = path.Substring(3);
                    protectedSegments = 0;
                }
                else return string.Empty;

                var segments = new List<string>();
                foreach (string segment in path.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    if (segment == ".") continue;
                    if (segment == "..")
                    {
                        if (segments.Count <= protectedSegments) return string.Empty;
                        segments.RemoveAt(segments.Count - 1);
                    }
                    else segments.Add(segment);
                }
                if (segments.Count < protectedSegments) return string.Empty;
                return prefix + string.Join("/", segments);
            }

            internal static ProjectMatch Find(IList<ProjectItem> projects, string modelPath)
            {
                string path = Normalize(modelPath);
                var match = new ProjectMatch();
                if (path.Length == 0) return match;
                int bestLength = -1;
                foreach (var project in projects)
                {
                    foreach (string root in project.Paths ?? new string[0])
                    {
                        string normalizedRoot = Normalize(root);
                        if (normalizedRoot.Length == 0) continue;
                        string folderPrefix = normalizedRoot.TrimEnd('/') + "/";
                        if (!string.Equals(path, normalizedRoot, StringComparison.OrdinalIgnoreCase) &&
                            !path.StartsWith(folderPrefix, StringComparison.OrdinalIgnoreCase)) continue;
                        if (normalizedRoot.Length > bestLength)
                        {
                            bestLength = normalizedRoot.Length;
                            match.Project = project;
                            match.RootPath = root;
                            match.IsAmbiguous = false;
                        }
                        else if (normalizedRoot.Length == bestLength && !ReferenceEquals(match.Project, project))
                            match.IsAmbiguous = true;
                    }
                }
                if (match.IsAmbiguous)
                {
                    match.Project = null;
                    match.RootPath = null;
                }
                return match;
            }
        }

        internal sealed class FamilyRequest
        {
            internal RequestKind Kind { get; set; }
            internal string ContextKey { get; set; }
            internal string TypeName { get; set; }
            internal InstallationConfiguration Configuration { get; set; }
            internal string FamilyPath { get; set; }
            internal string OutputPath { get; set; }
            internal bool HasAutomaticName { get; set; }
            internal bool IsDirty { get; set; }
            internal FamilyTypeItem TargetType { get; set; }
            internal string PersistedName { get; set; }
        }

        // Изменения модели проходят через Execute; слежение за активной моделью — через Idling.
        // Классы вложены здесь, чтобы не расширять файловую структуру плагина.
        internal sealed class FamilyRequestHandler : IExternalEventHandler
        {
            private static readonly Dictionary<string, string> SavedFamilyPaths = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            private readonly List<string> _automaticReadDirectories = new List<string>();
            private string _activeContextKey;

            private static string ContextKey(Document document, ProjectItem project)
            {
                if (project != null) return "project:" + project.Name + "|" + project.Stage + "|" + project.Code;
                if (document == null || !document.IsValidObject) return "no-document";
                string path = GetModelPaths(document).FirstOrDefault(p => !string.IsNullOrWhiteSpace(p));
                if (string.IsNullOrWhiteSpace(path)) path = document.PathName;
                return !string.IsNullOrWhiteSpace(path) ? "document:" + path : "unsaved:" + document.GetHashCode();
            }

            private bool UpdateAutomaticFamily(UIApplication app, Document document, ProjectItem project, bool force = false)
            {
                string key = ContextKey(document, project);
                _activeContextKey = key;
                if (_owner.SwitchFamilyContext(key) && !force) return true;
                _owner.SetBusy(true);
                try
                {
                    FamilyPackage package = null;
                    string path = ProjectFamilyStorage.GetPath(project);
                    if (path != null && !EnsureProjectFamily(path)) return false;
                    string remembered;
                    if (path == null && SavedFamilyPaths.TryGetValue(key, out remembered) && File.Exists(remembered)) path = remembered;
                    if (path != null)
                    {
                        var opened = FindOpenFamily(app, path);
                        package = opened == null ? ReadFamilyPackage(app, path, true) : ReadOpenFamilyPackage(app, opened, path);
                    }
                    else if (document != null && document.IsValidObject && document.IsFamilyDocument
                        && document.FamilyManager.get_Parameter("Секции_Промежуточные_Количество") != null)
                    {
                        if (string.IsNullOrWhiteSpace(document.PathName) || !File.Exists(document.PathName))
                            throw new InvalidOperationException("Сначала сохраните открытое семейство в Revit. После этого его типы можно прочитать автоматически.");
                        bool library = SamePath(document.PathName, ResolveSourcePath());
                        package = library ? ReadFamilyPackage(document, null, false) : ReadOpenFamilyPackage(app, document, document.PathName);
                    }
                    else if (document != null && document.IsValidObject && !document.IsFamilyDocument)
                    {
                        string expected = Path.GetFileNameWithoutExtension(VentilationSettingsConfiguratorMain.MakeFileName(project?.Code));
                        var families = new FilteredElementCollector(document).OfClass(typeof(Family)).Cast<Family>()
                            .Where(f => f.IsEditable && (f.Name.StartsWith("550_Универсальная установка_Одноуровневая_(", StringComparison.OrdinalIgnoreCase)
                                || f.Name.StartsWith("Вентустановка_", StringComparison.OrdinalIgnoreCase))).ToList();
                        var exact = families.Where(f => string.Equals(f.Name, expected, StringComparison.OrdinalIgnoreCase)).ToList();
                        if (exact.Count == 1) families = exact;
                        if (families.Count > 1) throw new InvalidOperationException("В проекте несколько семейств вентустановок. Для однозначной автоматической загрузки нужно задать папку проектного семейства.");
                        if (families.Count == 1) package = ReadProjectFamily(document, families[0]);
                    }
                    if (package == null) package = ReadFamilyPackage(app, ResolveSourcePath(), false);
                    _owner.SetFamily(package);
                    return true;
                }
                catch (Exception ex)
                {
                    _owner.SetOperationError(OperationError.Create("Автоматическое чтение типов семейства", ex, null));
                    return false;
                }
                finally { _owner.SetBusy(false); }
            }

            private bool EnsureProjectFamily(string path)
            {
                bool ready = ProjectFamilyStorage.EnsureFile(ResolveSourcePath(), path, (source, destination) =>
                {
                    var dialog = new TaskDialog(PluginName)
                    {
                        MainInstruction = "Семейство проекта не найдено. Создать его?",
                        MainContent = "Файл проекта:\n" + destination
                            + "\n\nБудет создана папка, если её нет, и скопировано исходное семейство:\n" + source,
                        CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No,
                        DefaultButton = TaskDialogResult.No
                    };
                    return dialog.Show() == TaskDialogResult.Yes;
                });
                if (!ready) _owner.SetStatus("Создание проектного семейства отменено.", false);
                return ready;
            }

            private FamilyPackage ReadProjectFamily(Document project, Family family)
            {
                if (project.IsModifiable) throw new InvalidOperationException("Завершите текущую команду Revit перед чтением семейства.");
                string directory = Path.Combine(Path.GetTempPath(), "KPLN_Ventilation_Project_" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(directory); _automaticReadDirectories.Add(directory);
                Document familyDocument = null;
                try
                {
                    familyDocument = project.EditFamily(family);
                    string snapshot = Path.Combine(directory, family.Name + ".rfa");
                    using (var options = new SaveAsOptions { OverwriteExistingFile = false, MaximumBackups = 1 }) familyDocument.SaveAs(snapshot, options);
                    var package = ReadFamilyPackage(familyDocument, null, true);
                    package.SourcePath = snapshot;
                    foreach (var item in package.Types) item.SourcePath = snapshot;
                    return package;
                }
                finally
                {
                    if (familyDocument != null && familyDocument.IsValidObject && !familyDocument.Close(false))
                        throw new InvalidOperationException("Не удалось закрыть временный документ семейства проекта.");
                }
            }

            private void ClearAutomaticSnapshots()
            {
                foreach (string directory in _automaticReadDirectories)
                {
                    try { if (Directory.Exists(directory)) Directory.Delete(directory, true); }
                    catch (IOException) { }
                    catch (UnauthorizedAccessException) { }
                }
                _automaticReadDirectories.Clear();
            }

            private readonly VentilationSettingsConfiguratorMain _owner;
            private UIApplication _application;
            private bool _stopTrackingRequested;
            private IList<ProjectItem> _projects = new List<ProjectItem>();
            private string _lastModelKey;
            internal FamilyRequest PendingRequest { get; set; }

            internal FamilyRequestHandler(VentilationSettingsConfiguratorMain owner)
            {
                _owner = owner;
            }

            public string GetName() { return PluginName; }

            internal void StartTracking(UIApplication app)
            {
                _application = app;
                app.Idling += OnIdling;
                RefreshProjects(app);
            }

            // Можно вызывать из Window.Closed: здесь нет обращений к Revit API.
            internal void RequestStopTracking()
            {
                _stopTrackingRequested = true;
                PendingRequest = null;
            }

            private void OnIdling(object sender, IdlingEventArgs args)
            {
                // Отписка от событий Revit тоже требует API-контекста.
                // Выполняем её здесь, прежде чем читать состояние уже закрытого окна.
                if (_stopTrackingRequested)
                {
                    if (_application != null)
                    {
                        _application.Idling -= OnIdling;
                        _application = null;
                    }
                    ClearAutomaticSnapshots();
                    return;
                }
                if (_owner.IsBusy || _application == null) return;
                TryUpdateActiveProject(_application, false);
            }

            private void TryUpdateActiveProject(UIApplication app, bool force)
            {
                string filePath = string.Empty;
                try
                {
                    Document document = app.ActiveUIDocument == null ? null : app.ActiveUIDocument.Document;
                    // Читаем фактический путь отдельно от сопоставления проекта, в том числе для RFA.
                    if (document != null && document.IsValidObject)
                        filePath = document.PathName ?? string.Empty;
                    UpdateActiveProject(document, filePath, force);
                }
                catch (Exception ex)
                {
                    // Ошибка чтения пути не должна оставлять в заголовке предыдущий проект.
                    string key = "Ошибка определения проекта: " + ex.Message;
                    if (string.Equals(_lastModelKey, key, StringComparison.Ordinal)) return;
                    _lastModelKey = key;
                    _owner.SetProject(new ProjectMatch());
                    _owner.SetStatus(key, true);
                }
            }

            private void RefreshProjects(UIApplication app)
            {
                string databaseError = null;
                try { _projects = ProjectDatabase.Read(ProjectDatabasePath); }
                catch (Exception ex)
                {
                    _projects = new List<ProjectItem>();
                    databaseError = ex.Message;
                }
                TryUpdateActiveProject(app, true);
                if (databaseError != null && !_owner.HasOperationError) _owner.SetStatus(databaseError, true);
            }

            private static string[] GetModelPaths(Document document)
            {
                if (document == null || !document.IsValidObject || document.IsFamilyDocument || document.IsLinked)
                    return new[] { string.Empty, string.Empty };
                string centralPath = string.Empty;
                if (document.IsWorkshared)
                {
                    try
                    {
                        using (ModelPath modelPath = document.GetWorksharingCentralModelPath())
                            if (modelPath != null)
                                centralPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath);
                    }
                    catch (Autodesk.Revit.Exceptions.InvalidOperationException) { }
                }
                return new[] { centralPath, document.PathName ?? string.Empty };
            }

            private ProjectMatch MatchModel(string[] paths)
            {
                // Центральный путь приоритетен. К локальному переходим только при отсутствии совпадений.
                var central = ProjectPathMatcher.Find(_projects, paths[0]);
                return central.Project != null || central.IsAmbiguous
                    ? central : ProjectPathMatcher.Find(_projects, paths[1]);
            }

            private ProjectMatch MatchDocument(Document document)
            {
                // Открытое проектное RFA сохраняет привязку по своей папке и аббревиатуре.
                if (document != null && document.IsValidObject && document.IsFamilyDocument)
                    return ProjectFamilyStorage.Find(_projects, document.PathName);
                return MatchModel(GetModelPaths(document));
            }

            private void UpdateActiveProject(Document document, string filePath, bool force)
            {
                string[] paths = GetModelPaths(document);
                string displayPath = !string.IsNullOrWhiteSpace(filePath) ? filePath
                    : !string.IsNullOrWhiteSpace(paths[0]) ? paths[0]
                    : document == null || !document.IsValidObject ? "Нет открытого файла" : "Файл не сохранён";
                string key = displayPath + "\n" + string.Join("\n", paths) + "\n" + (document == null ? "" : document.GetHashCode().ToString());
                if (!force && string.Equals(_lastModelKey, key, StringComparison.Ordinal)) return;
                _lastModelKey = key;
                var match = MatchDocument(document);
                _owner.SetProject(match);
                UpdateAutomaticFamily(_application, document, match.Project);
            }

            private static IList<OpenDocumentItem> GetOpenProjects(UIApplication app)
            {
                var items = new List<OpenDocumentItem>();
                foreach (Document doc in app.Application.Documents)
                {
                    if (!doc.IsValidObject || doc.IsFamilyDocument || doc.IsLinked || doc.IsReadOnly)
                        continue;
                    items.Add(new OpenDocumentItem
                    {
                        Document = doc,
                        DisplayName = doc.Title,
                        LoadTargetName = doc.Title + (string.IsNullOrWhiteSpace(doc.PathName)
                            ? " — не сохранён" : " — " + doc.PathName)
                    });
                }
                return items.OrderBy(p => p.DisplayName).ToList();
            }

            public void Execute(UIApplication app)
            {
                var request = PendingRequest;
                PendingRequest = null;
                string savedPath = null;
                string stage = "Проверка настроек";
                string diagnostics = null;
                try
                {
                    if (request == null) return;
                    if (request.Kind == RequestKind.LoadSectionCatalog)
                    {
                        stage = "Чтение состава исходного семейства";
                        Document active = app.ActiveUIDocument == null ? null : app.ActiveUIDocument.Document;
                        if (UpdateAutomaticFamily(app, active, MatchDocument(active).Project, true)) _owner.CreateTypeIfReady();
                        return;
                    }
                    Document activeDocument = app.ActiveUIDocument == null ? null : app.ActiveUIDocument.Document;
                    ProjectItem project = MatchDocument(activeDocument).Project;
                    // Между нажатием кнопки и ExternalEvent пользователь мог сменить документ.
                    if (!string.Equals(request.ContextKey, ContextKey(activeDocument, project), StringComparison.OrdinalIgnoreCase))
                    {
                        TryUpdateActiveProject(app, true);
                        throw new InvalidOperationException("Активный проект изменился. Выберите нужный тип в текущем проекте и повторите действие.");
                    }
                    bool deleting = request.Kind == RequestKind.DeleteType;
                    string typeName = deleting ? request.PersistedName : ValidateTypeName(request.TypeName);
                    if (!deleting)
                    {
                        if (request.Configuration == null) throw new InvalidOperationException("Сначала создайте тип.");
                        request.Configuration.ValidateForOutput();
                    }
                    string configuredPath = ProjectFamilyStorage.GetPath(project);
                    var action = FamilyActionPlan.Make(request.Kind, request.OutputPath, configuredPath, request.PersistedName, request.IsDirty);

                    // Открытие и загрузка не заходят в сохранение и не меняют RFA на диске.
                    if (!action.WritesFamily)
                    {
                        string path = action.Path;
                        if (!File.Exists(path)) throw new FileNotFoundException("Сохранённое семейство недоступно. Проверьте путь к файлу.", path);
                        if (request.Kind == RequestKind.OpenInRevit)
                        {
                            stage = "Открытие сохранённой установки в Revit";
                            OpenSavedFamily(app, path, request.PersistedName, value => diagnostics = value);
                            _owner.SetStatus("Вент. установка открыта с типом «" + request.PersistedName + "».", false);
                        }
                        else
                        {
                            stage = "Выбор проекта для загрузки";
                            var projects = GetOpenProjects(app);
                            if (projects.Count == 0) throw new InvalidOperationException("Нет доступных проектов. Откройте проект в Revit и повторите загрузку.");
                            var chosen = _owner.ChooseLoadProject(projects, activeDocument);
                            if (chosen == null) { _owner.SetStatus("Выбор проекта отменён.", false); return; }
                            var target = GetTargetDocument(app, chosen);
                            stage = "Загрузка сохранённого типа в проект";
                            _owner.SetStatus(LoadSavedType(app, target, path, request.PersistedName), false);
                        }
                        return;
                    }

                    string sourcePath = ResolveSourcePath();
                    stage = "Определение файла для сохранения";
                    string outputPath = action.Path;
                    if (action.NeedsPath) outputPath = _owner.ChooseOutputPath(project?.Code, null);
                    if (outputPath == null) { _owner.SetStatus("Сохранение отменено. Файлы не изменены.", false); return; }
                    if (configuredPath != null && !EnsureProjectFamily(configuredPath)) return;
                    stage = "Проверка пути сохранения";
                    outputPath = ValidateOutputPath(app, sourcePath, outputPath, configuredPath);
                    var openFamily = FindOpenFamily(app, outputPath);
                    string inputPath = File.Exists(outputPath) ? outputPath : !string.IsNullOrWhiteSpace(request.FamilyPath) ? request.FamilyPath : sourcePath;
                    if (!File.Exists(inputPath)) throw new FileNotFoundException("Файл семейства недоступен. Проверьте путь и подключение к сети.", inputPath);
                    if (openFamily == null) ValidateDiskSource(app, inputPath);
                    bool freshLibraryCopy = SamePath(inputPath, sourcePath) || SamePath(inputPath, SourceFamilyPath) || SamePath(inputPath, LiteralSourceFamilyPath);
                    string persistedName = freshLibraryCopy && !deleting ? null : request.PersistedName;
                    FamilyPackage existing = null;
                    if (!freshLibraryCopy && File.Exists(outputPath))
                    {
                        stage = "Чтение существующих типов";
                        existing = openFamily == null ? ReadFamilyPackage(app, inputPath, true) : ReadFamilyPackage(openFamily, outputPath, true);
                        if (!deleting)
                        {
                            if (persistedName == null && request.HasAutomaticName)
                                typeName = _owner.AvailableAutomaticName(typeName, existing.Types.Select(t => t.Name));
                            _owner.MergeDiscoveredFamily(existing, typeName);
                            if (persistedName == null && existing.Types.Any(t => string.Equals(t.Name, typeName, StringComparison.OrdinalIgnoreCase)))
                                throw new InvalidOperationException("Тип «" + typeName + "» уже есть в файле и добавлен в список. Выберите его для редактирования или задайте другое имя новому типу.");
                        }
                    }
                    FamilyPackage savedFamily;
                    string warning;
                    bool unchanged = !deleting && !request.IsDirty && persistedName != null && SamePath(outputPath, request.OutputPath) && existing != null;
                    if (unchanged && existing != null && !existing.Types.Any(t => t.PersistedName == persistedName))
                        throw new InvalidOperationException("В семействе больше нет типа «" + persistedName + "».");
                    // Явное «Сохранить» исправляет и ранее созданный тип с длинами скрытых блоков из шаблона.
                    if (unchanged && existing.Types.Any(t => t.PersistedName == persistedName
                        && (t.Configuration == null || t.Configuration.NeedsInactiveLengthReset))) unchanged = false;
                    if (unchanged && openFamily == null)
                    {
                        savedFamily = existing; warning = string.Empty;
                    }
                    else if (openFamily != null)
                    {
                        savedFamily = SaveOpenFamily(openFamily, outputPath, typeName, persistedName, request.Configuration, deleting,
                            value => stage = value, value => diagnostics = value, !unchanged);
                        warning = string.Empty;
                    }
                    else warning = BuildAndSaveFamily(app, inputPath, outputPath, typeName, persistedName, freshLibraryCopy, request.Configuration,
                        value => stage = value, value => diagnostics = value, out savedFamily, deleting);
                    savedPath = outputPath;
                    if (_activeContextKey != null) SavedFamilyPaths[_activeContextKey] = outputPath;
                    if (deleting)
                    {
                        _owner.CompleteDeletedFamily(savedFamily, request.TargetType);
                        _owner.SetStatus("Тип «" + persistedName + "» удалён из семейства. Файл сохранён." + warning, false);
                    }
                    else
                    {
                        _owner.ApplyCalculatedDimensions(request.Configuration);
                        _owner.RememberSavedFamily(savedFamily, typeName);
                        _owner.SetStatus("Вент. установка сохранена с типом «" + typeName + "». " + savedPath + warning, false);
                    }
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException) { _owner.SetStatus("Операция отменена.", false); }
                catch (Exception ex)
                {
                    var error = OperationError.Create(stage, ex, savedPath, request != null && (request.Kind == RequestKind.Save || request.Kind == RequestKind.DeleteType));
                    _owner.SetOperationError(error);
                }
                finally { _owner.SetBusy(false); }
            }

            internal static string ValidateTypeName(string value)
            {
                string name = (value ?? string.Empty).Trim();
                if (name.Length == 0)
                    throw new ArgumentException("Введите имя нового типа.");
                if (string.Equals(name, BaseFamilyTypeName, StringComparison.OrdinalIgnoreCase))
                    throw new ArgumentException("Это имя базового типа. Введите другое имя для новой установки.");
                if (name.Any(char.IsControl) || name.IndexOfAny(new[] { '\\', ':', '{', '}', '[', ']', '|', ';', '<', '>', '?', '`', '~' }) >= 0)
                    throw new ArgumentException("Имя типа содержит недопустимые символы: \\ : { } [ ] | ; < > ? ` ~ или перевод строки.");
                return name;
            }

            private static Document GetTargetDocument(UIApplication app, OpenDocumentItem project)
            {
                Document target = project == null ? null : project.Document;
                if (target == null || !target.IsValidObject ||
                    !app.Application.Documents.Cast<Document>().Any(d => d.Equals(target)))
                    throw new InvalidOperationException("Откройте нужный проект в Revit и повторите действие.");
                if (target.IsFamilyDocument || target.IsLinked || target.IsReadOnly || target.IsModifiable)
                    throw new InvalidOperationException("Выбранный проект сейчас недоступен для загрузки. Завершите редактирование и повторите действие.");
                return target;
            }

            private static string ValidateOutputPath(UIApplication app, string sourcePath, string outputPath, string configuredPath)
            {
                string path = Path.GetFullPath(outputPath);
                if (!string.Equals(Path.GetExtension(path), ".rfa", StringComparison.OrdinalIgnoreCase))
                    throw new ArgumentException("Для копии необходимо выбрать файл с расширением .rfa.");
                if (SamePath(path, sourcePath) || SamePath(path, SourceFamilyPath) || SamePath(path, LiteralSourceFamilyPath))
                    throw new InvalidOperationException("Нельзя сохранять копию поверх исходного семейства. Выберите другой файл.");
                // В утверждённой папке проекта по условию используется исходное имя (Об).
                // Для произвольного пути сохраняем прежнюю защиту от выбора оригинала через UNC.
                if (!SamePath(path, configuredPath)
                    && string.Equals(Path.GetFileName(path), Path.GetFileName(sourcePath), StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Задайте копии другое имя файла, чтобы отличать её от исходного семейства.");
                if (!Directory.Exists(Path.GetDirectoryName(path)))
                    throw new DirectoryNotFoundException("Выбранная папка недоступна.");
                if (File.Exists(path) && (File.GetAttributes(path) & FileAttributes.ReadOnly) != 0)
                    throw new IOException(SamePath(path, configuredPath)
                        ? "Семейство проекта доступно только для чтения. Проверьте права на запись в файл проекта."
                        : "Файл копии доступен только для чтения. Выберите другое место сохранения.");
                return path;
            }

            private static bool SamePath(string a, string b)
            {
                // Открытые облачные проекты могут иметь PathName вида BIM 360://... .
                // Такие адреса нельзя передавать в Path.GetFullPath как пути Windows.
                if (!Path.IsPathRooted(a) || !Path.IsPathRooted(b)) return false;
                return string.Equals(Path.GetFullPath(a), Path.GetFullPath(b), StringComparison.OrdinalIgnoreCase);
            }

            private static string BuildAndSaveFamily(UIApplication app, string sourcePath, string outputPath, string typeName, string persistedName, bool freshLibraryCopy,
                InstallationConfiguration configuration, Action<string> setStage, Action<string> setDiagnostics, out FamilyPackage savedFamily, bool deleting = false)
            {
                savedFamily = null;
                // Подготовка на том же томе, что и результат: готовый RFA заменяет старый файл
                // только после успешной транзакции, SaveAs и Close. Исходник лишь читается File.Copy.
                string workDirectory = Path.Combine(Path.GetDirectoryName(outputPath),
                    ".KPLN_Ventilation_" + Guid.NewGuid().ToString("N"));
                string inputDirectory = Path.Combine(workDirectory, "input");
                string stagedDirectory = Path.Combine(workDirectory, "result");
                Document familyDoc = null;
                string warning = string.Empty;
                string diagnosticText = string.Empty;
                string outputHash = File.Exists(outputPath) ? FileHash(outputPath) : null;
                try
                {
                    setStage("Создание временной копии семейства");
                    Directory.CreateDirectory(inputDirectory);
                    Directory.CreateDirectory(stagedDirectory);
                    string inputPath = Path.Combine(inputDirectory, Path.GetFileName(sourcePath));
                    string stagedPath = Path.Combine(stagedDirectory, Path.GetFileName(outputPath));
                    File.Copy(sourcePath, inputPath, false);
                    File.SetAttributes(inputPath, File.GetAttributes(inputPath) & ~FileAttributes.ReadOnly);
                    setStage("Открытие временной копии семейства");
                    familyDoc = app.Application.OpenDocumentFile(inputPath);
                    if (!familyDoc.IsFamilyDocument)
                        throw new InvalidOperationException("Исходный файл не является редактируемым семейством Revit.");

                    ApplyFamilyMutation(familyDoc, sourcePath, typeName, persistedName, freshLibraryCopy, configuration, deleting, setStage,
                        value => { diagnosticText = value; setDiagnostics(value); });
                    setStage("Чтение сохранённых типов");
                    var package = ReadFamilyPackage(familyDoc, outputPath, true);

                    setStage("Сохранение подготовленного семейства");
                    using (var options = new SaveAsOptions())
                    {
                        options.OverwriteExistingFile = false;
                        options.MaximumBackups = 1;
                        familyDoc.SaveAs(stagedPath, options);
                    }
                    setStage("Закрытие подготовленного семейства");
                    if (!familyDoc.Close(false))
                        throw new InvalidOperationException("Не удалось закрыть подготовленную копию. Файл результата не перезаписан.");
                    familyDoc = null;

                    setStage("Запись итогового файла");
                    if (outputHash != (File.Exists(outputPath) ? FileHash(outputPath) : null))
                        throw new IOException("Файл семейства изменился во время сохранения. Результат не перезаписан; повторите загрузку файла.");
                    if (File.Exists(outputPath))
                    {
                        // Без небезопасного fallback Delete + Copy: если том не поддерживает Replace,
                        // операция завершится ошибкой, а существующий результат останется на месте.
                        File.Replace(stagedPath, outputPath, Path.Combine(workDirectory, "previous.rfa"));
                    }
                    else
                        File.Move(stagedPath, outputPath);
                    savedFamily = package;
                }
                catch
                {
                    // Только после ошибки сохранения и отката, никогда во время ввода в интерфейсе.
                    // Сохраняем первоначальную ошибку даже при неполной диагностике вложенных семейств.
                    if (!deleting && configuration != null && familyDoc != null && familyDoc.IsValidObject && !familyDoc.IsModifiable)
                    {
                        try { setDiagnostics(diagnosticText + "\n\n" + ReadNestedDiagnostics(familyDoc, configuration)); }
                        catch (Exception ex) { setDiagnostics(diagnosticText + "\nВложенные зависимости не прочитаны: " + OperationError.Message(ex)); }
                    }
                    throw;
                }
                finally
                {
                    bool canClean = true;
                    if (familyDoc != null && familyDoc.IsValidObject)
                    {
                        try { canClean = familyDoc.Close(false); }
                        catch { canClean = false; }
                    }
                    if (canClean && Directory.Exists(workDirectory))
                    {
                        try { Directory.Delete(workDirectory, true); }
                        catch (IOException) { warning = "\nНе удалось удалить временную папку: " + workDirectory; }
                        catch (UnauthorizedAccessException) { warning = "\nНе удалось удалить временную папку: " + workDirectory; }
                    }
                    else if (!canClean)
                        warning = "\nВременный документ не закрыт: " + workDirectory;
                }
                return warning;
            }

            private static void ApplyFamilyMutation(Document familyDoc, string sourcePath, string typeName, string persistedName, bool freshLibraryCopy,
                InstallationConfiguration configuration, bool deleting, Action<string> setStage, Action<string> setDiagnostics, Action afterCommit = null)
            {
                setStage("Чтение исходного состояния семейства");
                var manager = familyDoc.FamilyManager;
                var plan = deleting ? null : TypeWritePlan.Make(manager.Types.Cast<FamilyType>().Select(t => t.Name), typeName, persistedName, freshLibraryCopy);
                var initialState = FamilyDiagnosticSnapshot.Read(familyDoc);
                string diagnosticText = "Семейство: " + sourcePath + "\n" + (deleting ? "Удаление типа «" + persistedName + "»."
                    : "Режим: " + (plan.Create ? "новый тип средствами Revit" : "редактирование существующего типа")
                        + "\nИсходный тип: " + plan.SourceTypeName + "\nИтоговый тип: " + typeName + "\n\n" + RequestedValuesReport(configuration))
                    + "\n\n" + initialState.Report("До изменения семейства");
                Action<string> appendDiagnostics = text =>
                { diagnosticText += "\n\n" + text; setDiagnostics(diagnosticText); };
                setDiagnostics(diagnosticText);
                var sourceInputs = deleting ? null : CaptureTypeInputs(manager,
                    manager.Types.Cast<FamilyType>().Single(t => t.Name == plan.SourceTypeName));
                var preserved = manager.Types.Cast<FamilyType>()
                    .Where(t => deleting ? t.Name != persistedName : plan.Create || t.Name != persistedName)
                    .Select(t => new PreservedType { Type = t, Name = t.Name, Inputs = CaptureTypeInputs(manager, t) }).ToList();
                if (!deleting)
                    appendDiagnostics(ReadFamilyDependencies(familyDoc, manager.Types.Cast<FamilyType>().Single(t => t.Name == plan.SourceTypeName))
                        .Report(InstallationConfiguration.InterfaceParameterNames()));

                // Исходный тип нужен только для наследования параметров. Не подтверждаем
                // его геометрию отдельной транзакцией: Revit проверяет сразу итоговый тип.
                // Группа позволяет откатить единственную транзакцию, если проверка после Commit не прошла.
                using (var group = new TransactionGroup(familyDoc, deleting ? "Удалить тип вентиляционной установки" : "Сохранить тип вентиляционной установки"))
                {
                    if (group.Start() != TransactionStatus.Started) throw new InvalidOperationException("Не удалось начать изменение семейства.");
                    try
                    {
                        if (deleting)
                            RunFamilyStage(familyDoc, "Удаление типа", () => DeleteFamilyType(manager, persistedName), null, setStage, appendDiagnostics);
                        else
                        {
                            Action verifyOtherValves = null;
                            RunFamilyStage(familyDoc, "Копирование и настройка итогового типа",
                                () =>
                                {
                                    setStage("Выбор исходного типа и создание копии");
                                    PrepareFamilyType(manager, typeName, persistedName, freshLibraryCopy);
                                    VerifyTypeInputs(manager.CurrentType, sourceInputs, null, "Наследование исходных значений нарушено");
                                    RefreshDimensionModes(manager, configuration);
                                    setStage("Применение параметров к итоговому типу");
                                    ApplySectionCount(manager, configuration.IntermediateCount);
                                    verifyOtherValves = ApplyValveState(manager, configuration.HasValve);
                                    foreach (var flag in configuration.BooleanAssignments()) SetIntegerValue(manager, flag.Key, flag.Value);
                                    ApplySectionTypes(familyDoc, configuration);
                                    ApplyDimensions(manager, configuration);
                                    ApplyInfo(manager, configuration.Info);
                                },
                                () =>
                                {
                                    VerifyRequestedValues(familyDoc, configuration);
                                    verifyOtherValves?.Invoke();
                                    // Связи геометрии и соединителей могут пересчитать служебные размеры
                                    // даже без формулы у FamilyParameter. Для текущего типа это штатно:
                                    // проверяем записанные настройки выше, а не равенство скрытых значений.
                                    // Сохранность остальных типов проверяется отдельно ниже.
                                }, setStage, appendDiagnostics);
                            appendDiagnostics("Итоговый тип подтверждён одной транзакцией. Промежуточные типы отдельно не подтверждались.");
                        }
                        appendDiagnostics(FamilyDiagnosticSnapshot.Read(familyDoc).CompareTo(initialState));
                        setStage("Проверка сохранности остальных типов");
                        foreach (var item in preserved)
                        {
                            if (!manager.Types.Cast<FamilyType>().Any(t => t.Name == item.Name))
                                throw new InvalidOperationException("Исчез исходный тип «" + item.Name + "». Изменения отменены.");
                            VerifyTypeInputs(item.Type, item.Inputs, null, "Изменился другой тип «" + item.Name + "»");
                        }
                        if (!deleting) ReadCalculatedDimensions(familyDoc, configuration);
                        setStage("Подтверждение всех изменений семейства");
                        if (group.Assimilate() != TransactionStatus.Committed)
                            throw new InvalidOperationException("Revit не подтвердил изменения семейства.");
                    }
                    finally
                    {
                        if (group.GetStatus() == TransactionStatus.Started) group.RollBack();
                    }
                }
                afterCommit?.Invoke();
            }

            private static void RunFamilyStage(Document document, string stage, Action change, Action verify, Action<string> setStage,
                Action<string> reportFailures = null)
            {
                setStage(stage);
                using (var transaction = new Transaction(document, stage))
                {
                    if (transaction.Start() != TransactionStatus.Started) throw new InvalidOperationException("Не удалось начать этап: " + stage + ".");
                    var failures = new TransactionFailures(ignoreEmptyBlockDuplicates: true);
                    var options = transaction.GetFailureHandlingOptions();
                    options.SetFailuresPreprocessor(failures); options.SetClearAfterRollback(true); options.SetForcedModalHandling(true);
                    transaction.SetFailureHandlingOptions(options);
                    try
                    {
                        change();
                        // Commit пересчитывает сразу итоговую конфигурацию. Отдельных Regenerate для шаблона и пустой копии нет.
                        setStage(stage + " — подтверждение итогового состояния");
                        if (transaction.Commit() != TransactionStatus.Committed)
                            throw new InvalidOperationException(failures.DescribeErrors("Revit отменил этап «" + stage + "». Изменения семейства отменены."));
                    }
                    finally
                    {
                        if (transaction.GetStatus() == TransactionStatus.Started) transaction.RollBack();
                        if (failures.HasMessages) reportFailures?.Invoke(failures.Describe("Сообщения Revit при подтверждении этапа «" + stage + "»:"));
                    }
                    setStage(stage + " — проверка результата");
                    verify?.Invoke();
                }
            }

            private static Document FindOpenFamily(UIApplication app, string path)
            {
                var document = app.Application.Documents.Cast<Document>().FirstOrDefault(d => d.IsValidObject && SamePath(d.PathName, path));
                if (document != null && (!document.IsFamilyDocument || document.IsLinked))
                    throw new InvalidOperationException("По этому пути открыт документ, который нельзя редактировать как семейство.");
                return document;
            }

            private FamilyPackage SaveOpenFamily(Document document, string path, string typeName, string persistedName,
                InstallationConfiguration configuration, bool deleting, Action<string> setStage, Action<string> setDiagnostics, bool applyChanges = true)
            {
                if (document.IsReadOnly || document.IsModifiable) throw new InvalidOperationException("Завершите редактирование в Revit: открытое семейство сейчас недоступно для записи.");
                bool applied = false;
                try
                {
                    if (applyChanges)
                        ApplyFamilyMutation(document, path, typeName, persistedName, false, configuration, deleting, setStage, setDiagnostics,
                            () => { applied = true; if (!deleting) _owner.RememberOpenTypeIdentity(path, typeName); });
                    setStage("Чтение изменённых типов открытого семейства");
                    var family = ReadFamilyPackage(document, path, true);
                    setStage("Сохранение открытого семейства");
                    document.Save(); // Transaction и TransactionGroup к этому моменту закрыты.
                    return family;
                }
                catch (Exception ex)
                {
                    if (!applied) throw;
                    throw new InvalidOperationException("Изменения применены в открытом семействе, но сохранить файл не удалось. "
                        + "Они остаются в документе Revit; повторите сохранение. " + OperationError.Message(ex), ex);
                }
            }

            private static bool DeleteFamilyType(FamilyManager manager, string typeName)
            {
                if (string.IsNullOrWhiteSpace(typeName) || string.Equals(typeName, BaseFamilyTypeName, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Шаблонный тип удалять нельзя.");
                var types = manager.Types.Cast<FamilyType>().ToList();
                var target = types.SingleOrDefault(t => string.Equals(t.Name, typeName, StringComparison.Ordinal));
                if (target == null) return false; // Повтор после неудачного Save открытого документа: тип уже удалён в памяти.
                if (types.Count <= 1) throw new InvalidOperationException("Нельзя удалить последний тип семейства.");
                manager.CurrentType = target;
                manager.DeleteCurrentType();
                return true;
            }

            private static void OpenSavedFamily(UIApplication app, string path, string typeName, Action<string> setDiagnostics)
            {
                // OpenAndActivateDocument активирует и уже открытый документ; повторного Save здесь нет.
                var document = app.OpenAndActivateDocument(path).Document;
                if (!document.IsFamilyDocument) throw new InvalidOperationException("Сохранённый файл не является семейством.");
                var manager = document.FamilyManager;
                var initialState = FamilyDiagnosticSnapshot.Read(document);
                string details = "Открытие сохранённого семейства: " + path + "\nЗапрошенный тип: " + typeName
                    + "\n\n" + initialState.Report("После открытия, до выбора типа");
                Action<string> appendDiagnostics = text => { details += "\n\n" + text; setDiagnostics(details); };
                setDiagnostics(details);
                var type = manager.Types.Cast<FamilyType>().SingleOrDefault(t => t.Name == typeName);
                if (type == null) throw new InvalidOperationException("В открытом семействе нет типа «" + typeName + "».");
                if (manager.CurrentType != null && manager.CurrentType.Name == typeName)
                {
                    appendDiagnostics("Запрошенный тип уже текущий. Транзакция выбора типа не выполнялась.");
                    return;
                }
                RunFamilyStage(document, "Показать выбранный тип установки", () => manager.CurrentType = type, null, stage => { }, appendDiagnostics);
                appendDiagnostics(FamilyDiagnosticSnapshot.Read(document).CompareTo(initialState));
            }

            private static string LoadSavedType(UIApplication app, Document target, string path, string typeName)
            {
                if (FindOpenFamily(app, path) == null) return LoadType(target, path, typeName);
                // Читаем сохранённые байты; несохранённые изменения открытого RFA не переносятся в проект.
                string directory = Path.Combine(Path.GetTempPath(), "KPLN_Ventilation_Load_" + Guid.NewGuid().ToString("N"));
                try
                {
                    Directory.CreateDirectory(directory);
                    string copy = Path.Combine(directory, Path.GetFileName(path));
                    File.Copy(path, copy, false);
                    return LoadType(target, copy, typeName);
                }
                finally
                {
                    try { if (Directory.Exists(directory)) Directory.Delete(directory, true); }
                    catch (IOException) { }
                    catch (UnauthorizedAccessException) { }
                }
            }

            private static string FileHash(string path)
            {
                using (var stream = File.OpenRead(path))
                using (var hash = System.Security.Cryptography.SHA256.Create())
                    return Convert.ToBase64String(hash.ComputeHash(stream));
            }

            private static TypeWritePlan PrepareFamilyType(FamilyManager manager, string typeName, string persistedName, bool freshLibraryCopy)
            {
                var plan = TypeWritePlan.Make(manager.Types.Cast<FamilyType>().Select(t => t.Name), typeName, persistedName, freshLibraryCopy);
                var source = SelectFamilyType(manager, plan.SourceTypeName);
                if (!plan.Create)
                {
                    if (source.Name != typeName) manager.RenameCurrentType(typeName);
                    return plan;
                }
                var inherited = CaptureTypeInputs(manager, source);
                manager.NewType(typeName);
                // Только проверка штатного наследования. Ни служебные, ни пользовательские
                // значения нового типа вручную из шаблона не пересобираем.
                VerifyTypeInputs(manager.CurrentType, inherited, null, "Новый тип не унаследовал значения исходного типа «" + source.Name + "»");
                return plan;
            }

            private static FamilyType SelectFamilyType(FamilyManager manager, string name)
            {
                var type = manager.Types.Cast<FamilyType>().Single(t => t.Name == name);
                if (manager.CurrentType == null || manager.CurrentType.Name != name) manager.CurrentType = type;
                return type;
            }

            private sealed class TypeInput
            {
                internal FamilyParameter Parameter;
                internal object Value;
            }
            private sealed class PreservedType
            {
                internal FamilyType Type;
                internal string Name;
                internal List<TypeInput> Inputs;
            }
            private static object ReadTypeInput(FamilyType type, FamilyParameter parameter)
            {
                switch (parameter.StorageType)
                {
                    case StorageType.Double: return type.AsDouble(parameter);
                    case StorageType.Integer: return type.AsInteger(parameter);
                    case StorageType.String: return type.AsString(parameter) ?? string.Empty;
                    case StorageType.ElementId:
                        var id = type.AsElementId(parameter);
                        return id == null ? (object)null : Common.IDHelper.ElIdValue(id);
                    default: return null;
                }
            }
            private static List<TypeInput> CaptureTypeInputs(FamilyManager manager, FamilyType type)
            {
                var result = new List<TypeInput>();
                foreach (FamilyParameter parameter in manager.Parameters)
                    if (!parameter.IsDeterminedByFormula && !parameter.IsReadOnly && !parameter.IsReporting)
                        result.Add(new TypeInput { Parameter = parameter, Value = ReadTypeInput(type, parameter) });
                return result;
            }
            private static void VerifyTypeInputs(FamilyType type, IEnumerable<TypeInput> inputs, ISet<string> allowed, string message)
            {
                var changed = new List<string>();
                foreach (var input in inputs)
                {
                    if (allowed != null && allowed.Contains(FamilyParameterNames.Canonical(input.Parameter.Definition.Name))) continue;
                    object current = ReadTypeInput(type, input.Parameter);
                    bool same = input.Value is double && current is double ? Math.Abs((double)input.Value - (double)current) < 1e-9 : Equals(input.Value, current);
                    if (!same) changed.Add(input.Parameter.Definition.Name + ": «" + Convert.ToString(input.Value, CultureInfo.InvariantCulture)
                        + "» → «" + Convert.ToString(current, CultureInfo.InvariantCulture) + "»");
                }
                if (changed.Count > 0) throw new InvalidOperationException(message + ":\n" + string.Join("\n", changed)
                    + "\nЧисловые значения здесь приведены во внутренних единицах Revit. Изменения отменены.");
            }

            private static void EnsureWritable(FamilyParameter parameter)
            {
                if (parameter.IsDeterminedByFormula)
                    throw new InvalidOperationException("Параметр «" + parameter.Definition.Name + "» вычисляется по формуле: " + parameter.Formula);
                if (parameter.IsReadOnly || parameter.IsReporting)
                    throw new InvalidOperationException("Параметр «" + parameter.Definition.Name + "» доступен только для чтения.");
            }

            private static void SetIntegerValue(FamilyManager manager, string name, int value)
            {
                var parameter = manager.get_Parameter(name);
                if (parameter == null || parameter.StorageType != StorageType.Integer)
                    throw new InvalidOperationException("В семействе нет целочисленного параметра «" + name + "».");
                // Существующее значение не требует записи, в том числе при формуле или блокировке параметра.
                if (manager.CurrentType.AsInteger(parameter) == value) return;
                // Формулы пересчитает Revit после записи независимых входов; результат проверяем перед сохранением.
                if (parameter.IsDeterminedByFormula) return;
                EnsureWritable(parameter);
                try { manager.Set(parameter, value); }
                catch (Exception ex)
                { throw new InvalidOperationException("Параметр «" + name + "», значение " + value + ": " + OperationError.Message(ex), ex); }
            }

            private static bool? ReadBooleanValue(FamilyManager manager, FamilyType type, string name)
            {
                var parameter = FindFamilyParameter(manager, name);
                if (parameter == null) return null;
                if (parameter.StorageType != StorageType.Integer)
                    throw new InvalidOperationException("Параметр «" + name + "» должен быть логическим (Да/Нет).");
                int? value = type.AsInteger(parameter);
                if (value.HasValue && value != 0 && value != 1)
                    throw new InvalidOperationException("Некорректное логическое значение «" + name + "».");
                return value.HasValue ? (bool?)(value.Value == 1) : null;
            }

            private static void ApplySectionCount(FamilyManager manager, int sectionCount)
            {
                if (sectionCount < 1 || sectionCount > 9)
                    throw new ArgumentException("В исходном семействе доступны от 1 до 9 промежуточных секций.");
                SetIntegerValue(manager, "Секции_Промежуточные_Количество", sectionCount);
            }

            private static Action ApplyValveState(FamilyManager manager, bool enabled)
            {
                var parameter = manager.get_Parameter(ValveControl.ParameterName);
                if (parameter == null || parameter.StorageType != StorageType.Integer)
                    throw new InvalidOperationException("В семействе нет логического параметра «" + ValveControl.ParameterName + "».");
                int requested = enabled ? 1 : 0;
                if (manager.CurrentType.AsInteger(parameter) == requested) return null;
                if (!parameter.IsDeterminedByFormula)
                { SetIntegerValue(manager, ValveControl.ParameterName, requested); return null; }

                bool? constant = ValveControl.ConstantFormula(parameter.Formula);
                if (!constant.HasValue || parameter.IsReadOnly || parameter.IsReporting)
                    throw new InvalidOperationException("Нельзя переключить клапан: параметр задан зависимой формулой или доступен только для чтения.");
                var target = manager.CurrentType;
                var preserved = manager.Types.Cast<FamilyType>().Where(t => t.Name != target.Name)
                    .Select(t => new { Type = t, Value = t.AsInteger(parameter) ?? (constant.Value ? 1 : 0) }).ToList();
                // Постоянная формула действует на всё семейство. Снимаем только её и сохраняем
                // прежние значения других типов, включая шаблон. Пересчёт — в общем Commit.
                manager.SetFormula(parameter, null);
                try
                {
                    foreach (var item in preserved)
                    {
                        if (item.Type.AsInteger(parameter) == item.Value) continue;
                        manager.CurrentType = item.Type;
                        SetIntegerValue(manager, ValveControl.ParameterName, item.Value);
                    }
                }
                finally { manager.CurrentType = target; }
                SetIntegerValue(manager, ValveControl.ParameterName, requested);
                return () =>
                {
                    foreach (var item in preserved)
                        if (item.Type.AsInteger(parameter) != item.Value)
                            throw new InvalidOperationException("Изменилось состояние клапана у другого типа «" + item.Type.Name + "». Изменения отменены.");
                };
            }

            private static FamilyParameter FindFamilyParameter(FamilyManager manager, string name)
            {
                var matches = FamilyParameterNames.Aliases(name).Select(manager.get_Parameter).Where(p => p != null).ToList();
                if (matches.Count > 1) throw new InvalidOperationException("Неоднозначное соответствие параметра «" + name + "»: "
                    + string.Join(", ", matches.Select(p => p.Definition.Name)));
                return matches.FirstOrDefault();
            }

            private static FamilyParameter LengthParameter(FamilyManager manager, string name)
            {
                var parameter = FindFamilyParameter(manager, name);
                if (parameter == null || parameter.StorageType != StorageType.Double)
                    throw new InvalidOperationException("В семействе нет параметра размера «" + name + "».");
                return parameter;
            }

            private static double ReadLengthMm(FamilyManager manager, FamilyType type, string name)
            {
                double? value = type.AsDouble(LengthParameter(manager, name));
                if (!value.HasValue || !(FamilyParameterNames.IsOffset(name) ? DimensionValue.IsFinite(value.Value) : DimensionValue.IsPositive(value.Value)))
                    throw new InvalidOperationException("В исходном типе не задан корректный размер «" + name + "».");
                return Common.IDHelper.ConvertInternalToMm(value.Value);
            }

            private static double ReadSourceLengthMm(FamilyManager manager, FamilyType type, string name)
            {
                double? value = type.AsDouble(LengthParameter(manager, name));
                // Незаданные значения отображаются пустыми. Подставных размеров здесь нет.
                return value.HasValue && DimensionValue.IsFinite(value.Value)
                    ? Common.IDHelper.ConvertInternalToMm(value.Value) : double.NaN;
            }

            private static void SetLengthMm(FamilyManager manager, string name, double millimeters)
            {
                if (!(FamilyParameterNames.IsOffset(name) ? DimensionValue.IsFinite(millimeters) : DimensionValue.IsPositive(millimeters)))
                    throw new ArgumentException("«" + name + "»: " + (FamilyParameterNames.IsOffset(name) ? "введите числовое смещение." : "размер должен быть больше нуля."));
                var parameter = LengthParameter(manager, name);
                double value = Common.IDHelper.ConvertMmToInternal(millimeters);
                double? current = manager.CurrentType.AsDouble(parameter);
                if (current.HasValue && Math.Abs(current.Value - value) < 1e-9) return;
                if (parameter.IsDeterminedByFormula || parameter.IsReadOnly || parameter.IsReporting) return;
                EnsureWritable(parameter);
                try { manager.Set(parameter, value); }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("Параметр «" + name + "», размер "
                    + millimeters.ToString(CultureInfo.CurrentCulture) + " мм: " + OperationError.Message(ex), ex);
                }
            }

            private static void ApplyDimensions(FamilyManager manager, InstallationConfiguration configuration)
            {
                foreach (var dimension in configuration.DimensionAssignments())
                    SetLengthMm(manager, dimension.Key, dimension.Value);
            }

            private static FamilyParameter InfoParameter(FamilyManager manager, string name)
            {
                var parameter = FindFamilyParameter(manager, name);
                if (parameter != null && parameter.StorageType != StorageType.String)
                    throw new InvalidOperationException("Параметр «" + name + "» должен быть текстовым.");
                return parameter;
            }

            private static string ReadInfoValue(FamilyManager manager, FamilyType type, string name)
            {
                var parameter = InfoParameter(manager, name);
                return parameter == null ? "" : type.AsString(parameter) ?? "";
            }

            private static void ApplyInfo(FamilyManager manager, InstallationInfo info)
            {
                foreach (var pair in info.Assignments())
                {
                    var parameter = InfoParameter(manager, pair.Key);
                    if (parameter == null)
                    {
                        if (pair.Value.Length == 0) continue;
                        throw new InvalidOperationException("В семействе отсутствует текстовый параметр «" + pair.Key
                            + "». Добавьте соответствующий параметр в исходное семейство, чтобы сохранить это поле.");
                    }
                    if ((manager.CurrentType.AsString(parameter) ?? "") == pair.Value) continue;
                    EnsureWritable(parameter);
                    try { manager.Set(parameter, pair.Value); }
                    catch (Exception ex)
                    { throw new InvalidOperationException("Параметр «" + pair.Key + "»: " + OperationError.Message(ex), ex); }
                }
            }

            private static void RefreshDimensionModes(FamilyManager manager, InstallationConfiguration configuration)
            {
                configuration.RefreshDimensionModes(name =>
                {
                    var parameter = LengthParameter(manager, name);
                    return parameter.IsDeterminedByFormula || parameter.IsReadOnly || parameter.IsReporting;
                });
                configuration.RefreshBooleanModes(name =>
                {
                    var parameter = FindFamilyParameter(manager, name);
                    if (parameter == null || parameter.StorageType != StorageType.Integer)
                        throw new InvalidOperationException("В семействе нет логического параметра «" + name + "».");
                    return parameter.IsDeterminedByFormula || parameter.IsReadOnly || parameter.IsReporting;
                });
            }

            private static void ReadCalculatedDimensions(Document document, InstallationConfiguration configuration)
            {
                RefreshDimensionModes(document.FamilyManager, configuration);
                foreach (var dimension in configuration.Dimensions())
                    if (dimension.Value.IsCalculated)
                        dimension.Value.ReadResult(ReadLengthMm(document.FamilyManager, document.FamilyManager.CurrentType, dimension.Key));
                foreach (var flag in configuration.SharedBooleans)
                    if (flag.Value.IsCalculated) flag.Value.ReadResult(ReadBooleanValue(document.FamilyManager, document.FamilyManager.CurrentType, flag.Key));
            }

            private static string RequestedValuesReport(InstallationConfiguration configuration)
            {
                var lines = new List<string> { "Заданные значения:", "Секции_Промежуточные_Количество = " + configuration.IntermediateCount,
                    "Соединитель_Приточный_Клапан = " + (configuration.HasValve ? "Да" : "Нет") };
                lines.AddRange(configuration.BooleanAssignments().Select(p => p.Key + " = " + (p.Value != 0 ? "Да" : "Нет")));
                lines.AddRange(configuration.Info.Assignments().Select(p => p.Key + " = " + p.Value));
                lines.AddRange(configuration.SectionAssignments().Select(p => InstallationConfiguration.ParameterName(p.Key) + " = " + p.Value.DisplayName));
                lines.AddRange(configuration.DimensionAssignments().Select(p => (configuration.Catalog.Dependencies.Find(p.Key)?.Name ?? p.Key)
                    + " = " + p.Value.ToString(CultureInfo.CurrentCulture) + " мм"));
                configuration.RecalculateLocal();
                lines.Add(configuration.Geometry.Report());
                return string.Join("\n", lines);
            }

            private static void ValueMismatch(FamilyParameter parameter, string requested, string actual)
            {
                throw new InvalidOperationException("Параметр «" + parameter.Definition.Name + "»: задано «" + requested
                    + "», после пересчёта Revit получилось «" + actual + "»."
                    + (parameter.IsDeterminedByFormula ? " Формула: " + parameter.Formula : " Проверьте зависимости этого параметра."));
            }

            private static void VerifyRequestedValues(Document document, InstallationConfiguration configuration)
            {
                var manager = document.FamilyManager; var type = manager.CurrentType;
                foreach (var pair in configuration.Info.Assignments())
                {
                    string actual = ReadInfoValue(manager, type, pair.Key);
                    if (actual != pair.Value) ValueMismatch(InfoParameter(manager, pair.Key), pair.Value, actual);
                }
                var integers = new Dictionary<string, int> { { "Секции_Промежуточные_Количество", configuration.IntermediateCount },
                    { "Соединитель_Приточный_Клапан", configuration.HasValve ? 1 : 0 } };
                foreach (var flag in configuration.BooleanAssignments()) integers.Add(flag.Key, flag.Value);
                foreach (var pair in integers)
                {
                    var parameter = manager.get_Parameter(pair.Key); int? actual = type.AsInteger(parameter);
                    if (actual != pair.Value) ValueMismatch(parameter, pair.Value.ToString(), actual.HasValue ? actual.Value.ToString() : "не задано");
                }
                foreach (var pair in configuration.DimensionAssignments())
                {
                    var parameter = LengthParameter(manager, pair.Key); double? actual = type.AsDouble(parameter);
                    if (parameter.IsDeterminedByFormula || parameter.IsReadOnly || parameter.IsReporting) continue;
                    double requested = Common.IDHelper.ConvertMmToInternal(pair.Value);
                    if (!actual.HasValue || double.IsNaN(actual.Value) || Math.Abs(actual.Value - requested) > 1e-9)
                        ValueMismatch(parameter, pair.Value.ToString(CultureInfo.CurrentCulture) + " мм", actual.HasValue
                            ? Common.IDHelper.ConvertInternalToMm(actual.Value).ToString(CultureInfo.CurrentCulture) + " мм" : "не задано");
                }
                foreach (var pair in configuration.SectionAssignments())
                {
                    var parameter = SectionParameter(manager, pair.Key); var id = type.AsElementId(parameter);
                    var actual = id == null || id.Equals(ElementId.InvalidElementId) ? null : DescribeType(document, id);
                    if (actual == null || actual.Key != pair.Value.Key) ValueMismatch(parameter, pair.Value.DisplayName, actual?.DisplayName ?? "не задано");
                }
                // При сохранении сверяем перенос правил с Revit. Во время ввода API не вызывается.
                configuration.Geometry = LocalGeometry.Calculate(configuration);
                if (string.IsNullOrWhiteSpace(configuration.Geometry.Limitation))
                    foreach (var pair in configuration.Geometry.EndsMm)
                    {
                        var parameter = manager.get_Parameter("Системный_" + pair.Key + "_Блок_Длина");
                        if (parameter == null || !pair.Value.HasValue) continue;
                        double? actual = type.AsDouble(parameter);
                        if (!actual.HasValue || Math.Abs(Common.IDHelper.ConvertInternalToMm(actual.Value) - pair.Value.Value) > 0.001)
                            ValueMismatch(parameter, pair.Value.Value.ToString(CultureInfo.CurrentCulture) + " мм (локальный расчёт)",
                                actual.HasValue ? Common.IDHelper.ConvertInternalToMm(actual.Value).ToString(CultureInfo.CurrentCulture) + " мм" : "не задано");
                    }
            }

            private static string DependencyElement(Element element)
            {
                if (element == null) return "Элемент не найден";
                var instance = element as FamilyInstance;
                string name;
                try { name = instance?.Symbol == null ? element.Name : instance.Symbol.FamilyName + " : " + instance.Symbol.Name; }
                catch (Autodesk.Revit.Exceptions.InvalidOperationException) { name = element.GetType().Name; }
                return name + " [ID " + Common.IDHelper.ElIdValue(element.Id) + "]";
            }

            private static FamilyDependencies ReadFamilyDependencies(Document document, FamilyType sourceType = null, IEnumerable<string> watchedNames = null)
            {
                var result = new FamilyDependencies();
                var parameters = document.FamilyManager.Parameters.Cast<FamilyParameter>().ToList();
                sourceType = sourceType ?? document.FamilyManager.Types.Cast<FamilyType>().FirstOrDefault(t => t.Name == BaseFamilyTypeName)
                    ?? document.FamilyManager.CurrentType;
                var byElement = new Dictionary<long, HashSet<string>>();
                var editable = new HashSet<string>((watchedNames ?? InstallationConfiguration.InterfaceParameterNames()).Select(FamilyParameterNames.Canonical), StringComparer.Ordinal);
                foreach (var parameter in parameters)
                {
                    var item = new ParameterDependency
                    {
                        Name = parameter.Definition.Name,
                        Formula = parameter.Formula,
                        IsReadOnly = parameter.IsReadOnly,
                        IsReporting = parameter.IsReporting,
                        IsInstance = parameter.IsInstance
                    };
                    if (sourceType != null)
                    {
                        try
                        {
                            item.SourceValue = parameter.StorageType == StorageType.String ? sourceType.AsString(parameter) : sourceType.AsValueString(parameter);
                            if (parameter.StorageType == StorageType.Integer) item.IntegerValue = sourceType.AsInteger(parameter);
                            if (parameter.StorageType == StorageType.ElementId)
                            { var id = sourceType.AsElementId(parameter); item.SourceValue = id == null ? "не задано" : DependencyElement(document.GetElement(id)); }
                        }
                        catch (Exception ex) { item.SourceValue = "не прочитано: " + OperationError.Message(ex); }
                    }
                    result.Parameters.Add(item.Name, item);
                }
                result.ConnectFormulas();
                foreach (var parameter in parameters)
                {
                    string name = parameter.Definition.Name;
                    try
                    {
                        foreach (Parameter associated in parameter.AssociatedParameters)
                        {
                            var element = associated.Element;
                            result.Parameters[name].Links.Add("Связь: " + DependencyElement(element) + " → " + associated.Definition.Name);
                            if (element == null || !editable.Contains(FamilyParameterNames.Canonical(name))) continue;
                            long id = Common.IDHelper.ElIdValue(element.Id); HashSet<string> names;
                            if (!byElement.TryGetValue(id, out names)) byElement[id] = names = new HashSet<string>();
                            names.Add(name);
                        }
                    }
                    catch (Exception ex) { result.ReadWarnings.Add("Не удалось полностью прочитать связи «" + name + "»: " + OperationError.Message(ex)); }
                }
                foreach (Dimension dimension in new FilteredElementCollector(document).OfClass(typeof(Dimension)))
                {
                    try
                    {
                        var related = new HashSet<string>();
                        string labelStatus = "без метки";
                        // У выравниваний FamilyLabel недоступен: это не повод потерять сами привязки.
                        try { var label = dimension.FamilyLabel; if (label != null) { related.Add(label.Definition.Name); labelStatus = label.Definition.Name; } }
                        catch (Autodesk.Revit.Exceptions.InvalidOperationException) { labelStatus = "метка неприменима"; }
                        var references = dimension.References == null ? new List<Reference>() : dimension.References.Cast<Reference>().ToList();
                        foreach (long id in references.Select(r => Common.IDHelper.ElIdValue(r.ElementId)).Distinct())
                        { HashSet<string> names; if (byElement.TryGetValue(id, out names)) related.UnionWith(names); }
                        if (related.Count == 0) continue;
                        string locking;
                        try
                        {
                            locking = dimension.NumberOfSegments > 0
                                ? "цепочка; EQ=" + dimension.AreSegmentsEqual + "; замки=" + string.Join(",", dimension.Segments.Cast<DimensionSegment>().Select(s => s.IsLocked))
                                : "замок=" + dimension.IsLocked;
                        }
                        catch (Exception ex) { locking = "замки не прочитаны: " + OperationError.Message(ex); }
                        string description = "Размер / привязка [ID " + Common.IDHelper.ElIdValue(dimension.Id) + "] [" + locking + "; " + labelStatus + "] → "
                            + string.Join("; ", references.Select(r => ReferenceDescription(document, r)));
                        foreach (string name in related) if (result.Parameters.ContainsKey(name)) result.Parameters[name].Links.Add(description);
                    }
                    catch (Exception ex) { result.ReadWarnings.Add("Не удалось прочитать размер [ID " + Common.IDHelper.ElIdValue(dimension.Id) + "]: " + OperationError.Message(ex)); }
                }
                return result;
            }

            private static string ReferenceDescription(Document document, Reference reference)
            {
                var element = document.GetElement(reference.ElementId);
                string description = DependencyElement(element);
                try
                {
                    var instance = element as FamilyInstance;
                    if (instance != null) description += " / опора «" + instance.GetReferenceName(reference) + "»";
                    description += " / " + reference.ConvertToStableRepresentation(document);
                }
                catch (Exception ex) { description += " / опора не прочитана: " + OperationError.Message(ex); }
                return description;
            }

            private static string ReadNestedDiagnostics(Document document, InstallationConfiguration configuration)
            {
                var lines = new List<string> { "Вложенные семейства: параметры размеров, формулы и привязки (без изменения файлов)." };
                var requested = configuration.Blocks.Select(b => b.Type).ToList();
                var manager = document.FamilyManager;
                var sourceType = manager.CurrentType ?? manager.Types.Cast<FamilyType>().FirstOrDefault(t => t.Name == BaseFamilyTypeName);
                if (sourceType != null)
                    foreach (int slot in new[] { 1, 2, 12 })
                    {
                        var original = DescribeType(document, sourceType.AsElementId(SectionParameter(manager, slot)));
                        requested.Add(original);
                        lines.Add("Исходный тип «" + InstallationConfiguration.ParameterName(slot) + "»: " + original.DisplayName);
                    }
                string[] sizes = { "КП_Р_Длина", "КП_Р_Ширина", "КП_Р_Высота", "КП_Р_Смещение", "Видимые" };
                var families = new FilteredElementCollector(document).OfClass(typeof(Family)).Cast<Family>().ToList();
                foreach (var group in requested.GroupBy(c => c.FamilyName))
                {
                    var family = families.SingleOrDefault(f => f.Name == group.Key);
                    if (family == null || !family.IsEditable) { lines.Add("Недоступно для чтения: " + group.Key); continue; }
                    Document nested = null;
                    bool closed = true;
                    try
                    {
                        nested = document.EditFamily(family);
                        foreach (var choice in group.GroupBy(c => c.Key).Select(g => g.First()))
                        {
                            lines.Add("\nВложенный тип: " + choice.DisplayName);
                            var type = nested.FamilyManager.Types.Cast<FamilyType>().FirstOrDefault(t => t.Name == choice.TypeName);
                            if (type == null) { lines.Add("Тип не найден в редактируемом вложенном семействе."); continue; }
                            lines.Add(ReadFamilyDependencies(nested, type, sizes).Report(sizes));
                        }
                    }
                    catch (Exception ex) { lines.Add("Не удалось прочитать «" + group.Key + "»: " + OperationError.Message(ex)); }
                    finally
                    {
                        if (nested != null && nested.IsValidObject)
                            try { closed = nested.Close(false); }
                            catch (Exception ex) { closed = false; lines.Add("Закрытие вложенного документа: " + OperationError.Message(ex)); }
                    }
                    if (!closed) { lines.Add("Не удалось закрыть временный вложенный документ «" + group.Key + "»."); break; }
                }
                return string.Join("\n", lines);
            }

            private static FamilyParameter SectionParameter(FamilyManager manager, int slot)
            {
                string name = InstallationConfiguration.ParameterName(slot);
                var parameter = FindFamilyParameter(manager, name);
                if (parameter == null || parameter.StorageType != StorageType.ElementId)
                    throw new InvalidOperationException("В исходном семействе нет параметра состава «" + name + "».");
                return parameter;
            }

            private static SectionTypeChoice DescribeType(Document document, ElementId id)
            {
                var element = document.GetElement(id);
                var type = element as ElementType;
                if (type != null) return new SectionTypeChoice(type.FamilyName, type.Name);
                var nested = element as NestedFamilyTypeReference;
                if (nested != null) return new SectionTypeChoice(nested.FamilyName, nested.TypeName);
                throw new InvalidOperationException("Не удалось прочитать название вложенного типа оборудования.");
            }

            private static List<KeyValuePair<ElementId, SectionTypeChoice>> ReadTypeChoices(Document document, FamilyParameter parameter)
            {
                return document.OwnerFamily.GetFamilyTypeParameterValues(parameter.Id)
                    .Select(id => new KeyValuePair<ElementId, SectionTypeChoice>(id, DescribeType(document, id))).ToList();
            }

            private static void ValidateDiskSource(UIApplication app, string path)
            {
                foreach (Document open in app.Application.Documents)
                    if (open.IsValidObject && SamePath(open.PathName, path) && open.IsModified)
                        throw new InvalidOperationException("В Revit есть несохранённые изменения этого семейства. Сохраните их перед чтением файла в конфигураторе.");
            }

            private static FamilyPackage ReadFamilyPackage(UIApplication app, string path, bool includeTypes, bool allowOpenChanges = false)
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) throw new FileNotFoundException("Файл семейства недоступен.", path);
                if (!allowOpenChanges) ValidateDiskSource(app, path);
                string directory = Path.Combine(Path.GetTempPath(), "KPLN_Ventilation_Read_" + Guid.NewGuid().ToString("N"));
                Document document = null;
                try
                {
                    Directory.CreateDirectory(directory);
                    string input = Path.Combine(directory, Path.GetFileName(path));
                    File.Copy(path, input, false);
                    File.SetAttributes(input, File.GetAttributes(input) & ~FileAttributes.ReadOnly);
                    document = app.Application.OpenDocumentFile(input);
                    if (!document.IsFamilyDocument) throw new InvalidOperationException("Выбранный файл не является семейством Revit.");
                    bool library = SamePath(path, ResolveSourcePath()) || SamePath(path, SourceFamilyPath) || SamePath(path, LiteralSourceFamilyPath);
                    var result = ReadFamilyPackage(document, library ? null : path, includeTypes && !library);
                    if (!document.Close(false)) throw new InvalidOperationException("Не удалось закрыть временный документ после чтения типов.");
                    document = null;
                    return result;
                }
                finally
                {
                    bool closed = document == null || !document.IsValidObject;
                    if (!closed) { try { closed = document.Close(false); } catch { closed = false; } }
                    if (closed && Directory.Exists(directory))
                    {
                        try { Directory.Delete(directory, true); }
                        catch (IOException) { }
                        catch (UnauthorizedAccessException) { }
                    }
                }
            }

            private static FamilyPackage ReadOpenFamilyPackage(UIApplication app, Document document, string path)
            {
                var family = ReadFamilyPackage(document, path, true);
                if (!document.IsModified) return family;
                var saved = ReadFamilyPackage(app, path, true, true);
                foreach (var type in family.Types)
                {
                    var disk = saved.Types.FirstOrDefault(t => t.PersistedName == type.PersistedName);
                    if (type.Configuration == null || disk?.Configuration == null
                        || type.Configuration.Signature() != disk.Configuration.Signature()) type.RequireSave();
                }
                return family;
            }

            private static FamilyPackage ReadFamilyPackage(Document document, string path, bool includeTypes)
            {
                var manager = document.FamilyManager;
                var source = manager.Types.Cast<FamilyType>().SingleOrDefault(t => t.Name == BaseFamilyTypeName);
                if (source == null) throw new InvalidOperationException("В семействе нет исходного типа «" + BaseFamilyTypeName
                    + "». Нельзя прочитать начальные значения для его копии.");
                var result = new FamilyPackage { Path = path, SourcePath = path, SelectedTypeName = manager.CurrentType?.Name, Catalog = ReadSectionCatalog(document, source) };
                if (includeTypes)
                    foreach (var type in manager.Types.Cast<FamilyType>().Where(t => t.Name != BaseFamilyTypeName).OrderBy(t => t.Name, StringComparer.CurrentCultureIgnoreCase))
                    {
                        var item = new FamilyTypeItem { Name = type.Name, PersistedName = type.Name, SavedPath = path, SourcePath = path };
                        try
                        {
                            item.Configuration = ReadConfiguration(document, type, result.Catalog);
                            item.HasAutomaticName = type.Name == item.Configuration.Info.AutomaticTypeName;
                            item.MarkSaved();
                        }
                        catch (Exception ex) { item.LoadError = "Не удалось прочитать тип «" + type.Name + "»: " + OperationError.Message(ex); }
                        result.Types.Add(item);
                    }
                return result;
            }

            private static SectionCatalog ReadSectionCatalog(Document document, FamilyType sourceType, SectionCatalog choicesCatalog = null)
            {
                var manager = document.FamilyManager;
                var countParameter = manager.get_Parameter("Секции_Промежуточные_Количество");
                var valveParameter = manager.get_Parameter("Соединитель_Приточный_Клапан");
                if (countParameter == null || valveParameter == null || countParameter.StorageType != StorageType.Integer
                    || valveParameter.StorageType != StorageType.Integer)
                    throw new InvalidOperationException("Не найдены параметры количества секций и наличия клапана.");
                int? count = sourceType.AsInteger(countParameter), valve = sourceType.AsInteger(valveParameter);
                if (!count.HasValue || count < 1 || count > 9 || !valve.HasValue || (valve != 0 && valve != 1))
                    throw new InvalidOperationException("В типе «" + sourceType.Name + "» некорректное количество секций или состояние клапана.");
                var catalog = new SectionCatalog
                {
                    SourceTypeName = sourceType.Name,
                    InstallationWidthMm = ReadSourceLengthMm(manager, sourceType, "Установка_Ширина"),
                    InstallationHeightMm = ReadSourceLengthMm(manager, sourceType, "Установка_Высота"),
                    FrameHeightMm = ReadSourceLengthMm(manager, sourceType, "Основание_Рама_Высота"),
                    IntermediateCount = count.Value,
                    HasValve = valve.Value != 0
                };
                for (int slot = 1; slot <= 12; slot++)
                {
                    var parameter = SectionParameter(document.FamilyManager, slot);
                    var definition = new SectionDefinition { ParameterName = parameter.Definition.Name };
                    if (choicesCatalog == null)
                    {
                        var values = ReadTypeChoices(document, parameter);
                        definition.Choices = SectionDefinition.OrderChoices(values.Select(v => v.Value));
                        if (definition.Choices.Count == 0) throw new InvalidOperationException("Нет доступных типов для «" + parameter.Definition.Name + "».");
                        string category = document.GetElement(values[0].Key).Category?.Name;
                        definition.DisplayName = parameter.Definition.Name + (string.IsNullOrWhiteSpace(category) ? "" : "<" + category + ">");
                    }
                    else
                    {
                        definition.Choices = choicesCatalog[slot].Choices;
                        definition.DisplayName = choicesCatalog[slot].DisplayName;
                    }
                    var selectedId = sourceType.AsElementId(parameter);
                    if (selectedId != null && !selectedId.Equals(ElementId.InvalidElementId))
                    {
                        var selected = DescribeType(document, selectedId);
                        definition.SourceType = definition.Choices.SingleOrDefault(c => c.Key == selected.Key);
                        if (definition.SourceType == null) throw new InvalidOperationException("Исходный тип «" + selected.DisplayName
                            + "» недоступен для параметра «" + parameter.Definition.Name + "».");
                    }
                    string lengthName = InstallationConfiguration.DimensionParameterName(slot, "Длина");
                    definition.DefaultLengthMm = ReadSourceLengthMm(manager, sourceType, lengthName);
                    if (slot == 1 || slot == 12)
                    {
                        definition.DefaultWidthMm = ReadSourceLengthMm(manager, sourceType, InstallationConfiguration.DimensionParameterName(slot, "Ширина"));
                        definition.DefaultHeightMm = ReadSourceLengthMm(manager, sourceType, InstallationConfiguration.DimensionParameterName(slot, "Высота"));
                        definition.DefaultOffsetXMm = ReadSourceLengthMm(manager, sourceType, InstallationConfiguration.DimensionParameterName(slot, "Смещение по X"));
                        definition.DefaultOffsetYMm = ReadSourceLengthMm(manager, sourceType, InstallationConfiguration.DimensionParameterName(slot, "Смещение по Y"));
                    }
                    catalog.Add(slot, definition);
                }
                foreach (string name in SharedParameters.LengthNames)
                    catalog.SharedLengthsMm.Add(name, ReadSourceLengthMm(manager, sourceType, name));
                foreach (string name in SharedParameters.BooleanNames)
                    catalog.SharedFlags.Add(name, ReadBooleanValue(manager, sourceType, name));
                catalog.Dependencies = ReadFamilyDependencies(document, sourceType);
                return catalog;
            }

            private static InstallationConfiguration ReadConfiguration(Document document, FamilyType type, SectionCatalog catalog)
            {
                // Та же процедура чтения, но снимок значений принадлежит именно выбранному типу.
                var configuration = InstallationConfiguration.CreateDraft(ReadSectionCatalog(document, type, catalog));
                var manager = document.FamilyManager;
                configuration.Info = new InstallationInfo
                {
                    SystemName = ReadInfoValue(manager, type, InstallationInfo.SystemNameParameter),
                    Manufacturer = ReadInfoValue(manager, type, InstallationInfo.ManufacturerParameter),
                    Mark = ReadInfoValue(manager, type, InstallationInfo.MarkParameter),
                    Unit = ReadInfoValue(manager, type, InstallationInfo.UnitParameter),
                    Description = ReadInfoValue(manager, type, InstallationInfo.DescriptionParameter),
                    ProductCode = ReadInfoValue(manager, type, InstallationInfo.ProductCodeParameter),
                    MassText = ReadInfoValue(manager, type, InstallationInfo.MassTextParameter)
                };
                return configuration;
            }

            private static void ApplySectionTypes(Document document, InstallationConfiguration configuration)
            {
                FamilyManager manager = document.FamilyManager;
                foreach (var assignment in configuration.SectionAssignments())
                {
                    FamilyParameter parameter = SectionParameter(manager, assignment.Key);
                    // Идентификаторы разрешаем в создаваемой копии, а не переносим из закрытого документа каталога.
                    var matches = ReadTypeChoices(document, parameter).Where(v => v.Value.Key == assignment.Value.Key).ToList();
                    if (matches.Count != 1)
                        throw new InvalidOperationException("Не найден однозначный тип «" + assignment.Value.DisplayName + "» для «" + parameter.Definition.Name + "».");
                    var current = manager.CurrentType.AsElementId(parameter);
                    if (current != null && current.Equals(matches[0].Key)) continue;
                    if (parameter.IsDeterminedByFormula) continue;
                    EnsureWritable(parameter);
                    try { manager.Set(parameter, matches[0].Key); }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException("Параметр «" + parameter.Definition.Name + "», тип «"
                        + assignment.Value.DisplayName + "»: " + OperationError.Message(ex), ex);
                    }
                }
                // Размеры записываются отдельно. Текстовые параметры и механизмы здесь не меняем.
            }

            private sealed class FamilyDiagnosticSnapshot
            {
                private string _typeName;
                private Dictionary<string, string> _warnings;
                private HashSet<long> _instances;
                private readonly List<string> _readErrors = new List<string>();

                internal static FamilyDiagnosticSnapshot Read(Document document)
                {
                    var result = new FamilyDiagnosticSnapshot { _typeName = document.FamilyManager.CurrentType?.Name ?? "не выбран" };
                    try
                    {
                        var warnings = new Dictionary<string, string>(StringComparer.Ordinal);
                        foreach (var message in document.GetWarnings())
                        {
                            // Уже сохранённые предупреждения фильтруем так же, как сообщения транзакции.
                            if (TransactionFailures.IsEmptyBlockDuplicate(document, message.GetFailureDefinitionId(),
                                message.GetFailingElements(), message.GetAdditionalElements())) continue;
                            string failing = Ids(message.GetFailingElements());
                            string additional = Ids(message.GetAdditionalElements());
                            string key = message.GetFailureDefinitionId().Guid.ToString() + "|" + failing + "|" + additional;
                            warnings[key] = message.GetDescriptionText() + " [ID: " + failing + "]"
                                + (additional.Length == 0 ? "" : " [Связанные ID: " + additional + "]");
                        }
                        result._warnings = warnings;
                    }
                    catch (Exception ex) { result._readErrors.Add("Предупреждения не прочитаны: " + OperationError.Message(ex)); }
                    try
                    {
                        using (var collector = new FilteredElementCollector(document))
                            result._instances = new HashSet<long>(collector.OfClass(typeof(FamilyInstance)).WhereElementIsNotElementType()
                                .ToElementIds().Select(Common.IDHelper.ElIdValue));
                    }
                    catch (Exception ex) { result._readErrors.Add("Экземпляры не прочитаны: " + OperationError.Message(ex)); }
                    return result;
                }

                private static string Ids(IEnumerable<ElementId> ids)
                { return string.Join(", ", ids.Select(Common.IDHelper.ElIdValue).Distinct().OrderBy(id => id)); }

                internal string Report(string title)
                {
                    var lines = new List<string> { title + ":", "Текущий тип: " + _typeName,
                        "Вложенных экземпляров: " + (_instances == null ? "не прочитано" : _instances.Count.ToString()),
                        "Предупреждений для проверки: " + (_warnings == null ? "не прочитано" : _warnings.Count.ToString()) };
                    if (_warnings != null) lines.AddRange(_warnings.OrderBy(p => p.Key, StringComparer.Ordinal).Select(p => "  " + p.Value));
                    lines.AddRange(_readErrors);
                    return string.Join("\n", lines);
                }

                internal string CompareTo(FamilyDiagnosticSnapshot before)
                {
                    var lines = new List<string> { Report("После подтверждения итогового типа") };
                    if (_instances != null && before._instances != null)
                    {
                        var added = _instances.Except(before._instances).OrderBy(id => id).ToList();
                        var removed = before._instances.Except(_instances).OrderBy(id => id).ToList();
                        lines.Add("Добавленных вложенных экземпляров: " + added.Count
                            + (added.Count == 0 ? "" : " [ID: " + string.Join(", ", added) + "]"));
                        lines.Add("Удалённых вложенных экземпляров: " + removed.Count
                            + (removed.Count == 0 ? "" : " [ID: " + string.Join(", ", removed) + "]"));
                    }
                    if (_warnings != null && before._warnings != null)
                    {
                        var added = _warnings.Where(p => !before._warnings.ContainsKey(p.Key)).ToList();
                        int unchanged = _warnings.Keys.Count(before._warnings.ContainsKey);
                        int removed = before._warnings.Keys.Count(key => !_warnings.ContainsKey(key));
                        lines.Add("Предупреждения: совпадают с исходными — " + unchanged + "; появились — " + added.Count + "; исчезли — " + removed + ".");
                        lines.AddRange(added.Select(p => "  Появилось: " + p.Value));
                    }
                    lines.Add("Сравнение относится к текущим типам до и после операции; само по себе появление предупреждения не доказывает создание экземпляра.");
                    return string.Join("\n", lines);
                }
            }

            private sealed class TransactionFailures : IFailuresPreprocessor
            {
                private const string EmptyBlockFamilyName =
                    "550_Вл_Универсальная установка_Одноуровневая_Секция_Пустой блок_(Об)";
                private readonly bool _ignoreEmptyBlockDuplicates;
                private readonly List<string> _messages = new List<string>();
                private readonly List<string> _errors = new List<string>();
                internal TransactionFailures(bool ignoreEmptyBlockDuplicates = false)
                { _ignoreEmptyBlockDuplicates = ignoreEmptyBlockDuplicates; }
                internal bool HasMessages { get { return _messages.Count > 0; } }

                private bool IsIgnoredEmptyBlockWarning(FailuresAccessor accessor, FailureMessageAccessor message)
                {
                    return _ignoreEmptyBlockDuplicates && message.GetSeverity() == FailureSeverity.Warning
                        && IsEmptyBlockDuplicate(accessor.GetDocument(), message.GetFailureDefinitionId(),
                            message.GetFailingElementIds(), message.GetAdditionalElementIds());
                }

                internal static bool IsEmptyBlockDuplicate(Document document, FailureDefinitionId failureId,
                    IEnumerable<ElementId> failingIds, IEnumerable<ElementId> additionalIds)
                {
                    if (document == null || !document.IsFamilyDocument
                        || failureId.Guid != BuiltInFailures.OverlapFailures.DuplicateInstances.Guid) return false;
                    var failing = failingIds.Distinct().ToList();
                    if (failing.Count < 2) return false;
                    // Проверяем реальные элементы, а не локализованный текст или ID из одного отчёта.
                    return failing.Concat(additionalIds).Distinct().All(id =>
                    {
                        var symbol = (document.GetElement(id) as FamilyInstance)?.Symbol;
                        return symbol != null && string.Equals(symbol.FamilyName, EmptyBlockFamilyName, StringComparison.OrdinalIgnoreCase)
                            && string.Equals(symbol.Name, "Пустой блок", StringComparison.OrdinalIgnoreCase);
                    });
                }

                public FailureProcessingResult PreprocessFailures(FailuresAccessor accessor)
                {
                    bool error = false;
                    foreach (var message in accessor.GetFailureMessages().ToList())
                    {
                        if (IsIgnoredEmptyBlockWarning(accessor, message))
                        {
                            // Полностью игнорируем: не показываем окно и не добавляем сообщение в подробности.
                            accessor.DeleteWarning(message);
                            continue;
                        }
                        var severity = message.GetSeverity();
                        string description = "[" + severity + "; " + message.GetFailureDefinitionId().Guid + "] " + message.GetDescriptionText();
                        var ids = message.GetFailingElementIds().Concat(message.GetAdditionalElementIds())
                            .Select(Common.IDHelper.ElIdValue).Distinct().ToList();
                        if (ids.Count > 0) description += " [ID: " + string.Join(", ", ids) + "]";
                        if (!string.IsNullOrWhiteSpace(description) && !_messages.Contains(description)) _messages.Add(description);
                        if (severity == FailureSeverity.Error || severity == FailureSeverity.DocumentCorruption)
                        {
                            error = true;
                            if (!_errors.Contains(description)) _errors.Add(description);
                        }
                    }
                    // Ошибочный результат не сохраняем и не предлагаем снимать зависимости ради продолжения.
                    // Остальные предупреждения остаются в обработке Revit; в окно ошибки попадают только ошибки.
                    return error ? FailureProcessingResult.ProceedWithRollBack : FailureProcessingResult.Continue;
                }
                internal string Describe(string fallback)
                { return _messages.Count == 0 ? fallback : fallback + "\n" + string.Join("\n", _messages); }
                internal string DescribeErrors(string fallback)
                { return _errors.Count == 0 ? fallback : fallback + "\n" + string.Join("\n", _errors); }
            }

            private static string LoadType(Document target, string path, string typeName)
            {
                var loadOptions = new FamilyLoadOptions(Path.GetFileNameWithoutExtension(path));
                using (var transaction = new Transaction(target, "Загрузить тип вентиляционной установки"))
                {
                    if (transaction.Start() != TransactionStatus.Started)
                        throw new InvalidOperationException("Не удалось начать загрузку в проект.");
                    var reportedFailures = new TransactionFailures();
                    var failures = transaction.GetFailureHandlingOptions();
                    failures.SetFailuresPreprocessor(reportedFailures);
                    failures.SetClearAfterRollback(true);
                    failures.SetForcedModalHandling(true);
                    transaction.SetFailureHandlingOptions(failures);
                    FamilySymbol symbol;
                    bool loaded = target.LoadFamilySymbol(path, typeName, loadOptions, out symbol);
                    if (loadOptions.Cancelled)
                    {
                        transaction.RollBack();
                        return "Загрузка в проект отменена.";
                    }
                    if (!loaded || symbol == null)
                    {
                        transaction.RollBack();
                        return "Revit не выполнил загрузку типа в проект. Возможно, такая версия типа уже загружена.";
                    }
                    if (!symbol.IsActive) symbol.Activate();
                    if (transaction.Commit() != TransactionStatus.Committed)
                        throw new InvalidOperationException(reportedFailures.DescribeErrors("Revit отменил загрузку типа в проект."));
                }
                return "Тип «" + typeName + "» загружен в проект «" + target.Title + "». Экземпляр не размещён.";
            }
        }

        private sealed class FamilyLoadOptions : IFamilyLoadOptions
        {
            private readonly string _rootFamilyName;
            private bool? _overwriteAccepted;
            internal bool Cancelled { get; private set; }

            internal FamilyLoadOptions(string rootFamilyName) { _rootFamilyName = rootFamilyName; }

            public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
            {
                bool accepted = ConfirmOverwrite();
                overwriteParameterValues = accepted;
                return accepted;
            }

            private bool ConfirmOverwrite()
            {
                if (_overwriteAccepted.HasValue) return _overwriteAccepted.Value;
                var dialog = new TaskDialog(PluginName)
                {
                    MainInstruction = "В проекте уже есть это семейство. Обновить его?",
                    MainContent = "Будут обновлены определение семейства и значения типов из загружаемого семейства. Это может изменить уже размещённые экземпляры.",
                    CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No,
                    DefaultButton = TaskDialogResult.No
                };
                bool accepted = dialog.Show() == TaskDialogResult.Yes;
                _overwriteAccepted = accepted;
                Cancelled = !accepted;
                return accepted;
            }

            public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse,
                out FamilySource source, out bool overwriteParameterValues)
            {
                // Основное семейство само тоже может быть общим: для него разрешаем
                // обновление только после того же подтверждения, что и для обычного.
                if (string.Equals(sharedFamily.Name, _rootFamilyName, StringComparison.OrdinalIgnoreCase))
                {
                    bool accepted = ConfirmOverwrite();
                    source = accepted ? FamilySource.Family : FamilySource.Project;
                    overwriteParameterValues = accepted;
                    return accepted;
                }
                // Существующие общие вложенные семейства проекта сохраняем.
                source = FamilySource.Project;
                overwriteParameterValues = false;
                return true;
            }
        }
    }
}