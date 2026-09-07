using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Common;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Runtime.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WpfComboBox = System.Windows.Controls.ComboBox;

namespace KPLN_Tools.Forms
{
    internal enum TepDialogAction
    {
        None,
        Calculate,
        ShowLast,
        ImportSettings,
        ExportSettings,
        Clear,
        ExportCsv,
        ExportXlsx
    }

    [DataContract]
    internal sealed class TepSettings
    {
        [DataMember] public string Method { get; set; } = "New";
        [DataMember] public string RulesVersion { get; set; } = "1.0";
        [DataMember] public string BuildingType { get; set; } = "Residential";
        [DataMember] public string AreaBasis { get; set; } = "Rooms";
        [DataMember] public string Grouping { get; set; } = "Link";
        [DataMember] public string SelectedBuildings { get; set; } = string.Empty;
        [DataMember] public double ZeroElevationMeters { get; set; }
        [DataMember] public double GroundElevationMeters { get; set; }
        [DataMember] public string ExcludedLevelWords { get; set; } = "чердак,техподполье,подполье,кровля";
        [DataMember] public string BuildingParameter { get; set; } = "Корпус";
        [DataMember] public string BuildingTypeParameter { get; set; } = "Тип здания";
        [DataMember] public string FunctionParameter { get; set; } = "Назначение";
        [DataMember] public string ApartmentParameter { get; set; } = "Номер квартиры";
        [DataMember] public string SummerParameter { get; set; } = "Летнее помещение";
        [DataMember] public string SummerCoefficientParameter { get; set; } = "Коэффициент";
        [DataMember] public string ParkingParameter { get; set; } = "Машино-место";
        [DataMember] public string BuiltInParameter { get; set; } = "Встроенно-пристроенное";
        [DataMember] public string StandaloneParameter { get; set; } = "Отдельно стоящий";
        [DataMember] public string IncludeParameter { get; set; } = "Включить в ТЭП";
        [DataMember] public string ResidentialWords { get; set; } = "жил,квартир";
        [DataMember] public string PublicWords { get; set; } = "обществен,офис,торгов,детск,школ,медицин";
        [DataMember] public string TechnicalWords { get; set; } = "техническ,венткамера,щитовая,итп,насосная";
        [DataMember] public string SummerWords { get; set; } = "балкон,лоджия,терраса,веранда";
        [DataMember] public string ExcludedPublicWords { get; set; } = "коридор,тамбур,переход,лестнич,шахт,пандус,инженер";
        [DataMember] public string UndergroundWords { get; set; } = "подзем,подвал,цоколь,паркинг";
        [DataMember] public string FalseWords { get; set; } = "нет,false,0,исключить,не учитывать";
        [DataMember] public int RoundingDigits { get; set; } = 2;
        [DataMember] public bool DetailedWarnings { get; set; } = true;
        [DataMember] public bool ShowResults { get; set; } = true;
        [DataMember] public bool CreateSchedule { get; set; }
        [DataMember] public bool ExportCsv { get; set; }
        [DataMember] public bool ExportXlsx { get; set; }
        [DataMember] public bool CreateViews { get; set; }
        [DataMember] public bool Create3D { get; set; }
        [DataMember] public bool UpdateGraphics { get; set; } = true;
        [DataMember] public string ManualCorrections { get; set; } = string.Empty;
        [DataMember] public string Prefix { get; set; } = "ТЭП_";
        [DataMember] public List<string> SelectedIndicators { get; set; } = new List<string>();
        [DataMember] public List<string> IncludedSourceKeys { get; set; } = new List<string>();
        [DataMember] public List<string> GraphicsOnlySourceKeys { get; set; } = new List<string>();
        [DataMember] public List<TepFormulaRule> NewMethodRules { get; set; } = new List<TepFormulaRule>();
        [DataMember] public List<TepFormulaRule> OldMethodRules { get; set; } = new List<TepFormulaRule>();
    }

    [DataContract]
    internal sealed class TepFormulaRule
    {
        [DataMember] public string Number { get; set; }
        [DataMember] public string Name { get; set; }
        [DataMember] public string Expression { get; set; }
    }

    internal sealed class TepSourceOption
    {
        public string Key { get; set; }
        public string Name { get; set; }
        public string Kind { get; set; }
        public string Status { get; set; }
        public bool IsIncluded { get; set; }
        public bool GraphicsOnly { get; set; }
        public long LinkInstanceId { get; set; } = -1;
    }

    internal sealed class TepIndicatorOption
    {
        public bool IsSelected { get; set; }
        public string Number { get; set; }
        public string Name { get; set; }
        public string Unit { get; set; }
    }

    [DataContract]
    internal sealed class TepResultRow
    {
        [DataMember] public string Number { get; set; }
        [DataMember] public string Name { get; set; }
        [DataMember] public double RawValue { get; set; }
        [DataMember] public string DisplayValue { get; set; }
        [DataMember] public string Unit { get; set; }
        [DataMember] public string Building { get; set; }
        [DataMember] public string Source { get; set; }
        [DataMember] public string Method { get; set; }
        [DataMember] public string Status { get; set; }
        [DataMember] public string Comment { get; set; }
    }

    [DataContract]
    internal sealed class TepDetailRow
    {
        [DataMember] public string Indicator { get; set; }
        [DataMember] public string Source { get; set; }
        [DataMember] public string Building { get; set; }
        [DataMember] public string Level { get; set; }
        [DataMember] public string Category { get; set; }
        [DataMember] public string Element { get; set; }
        [DataMember] public long ElementId { get; set; }
        [DataMember] public long HostLinkId { get; set; }
        [DataMember] public string Function { get; set; }
        [DataMember] public string Apartment { get; set; }
        [DataMember] public double RawValue { get; set; }
        [DataMember] public string Unit { get; set; }
        [DataMember] public string Rule { get; set; }
    }

    [DataContract]
    internal sealed class TepWarningRow
    {
        [DataMember] public string Code { get; set; }
        [DataMember] public string Severity { get; set; }
        [DataMember] public string Description { get; set; }
        [DataMember] public string Indicator { get; set; }
        [DataMember] public string Source { get; set; }
        [DataMember] public string Building { get; set; }
        [DataMember] public string Category { get; set; }
        [DataMember] public long ElementId { get; set; }
        [DataMember] public long HostLinkId { get; set; }
        [DataMember] public string UserAction { get; set; }
        [DataMember] public string Accounting { get; set; }
    }

    [DataContract]
    internal sealed class TepCalculationOutput
    {
        [DataMember] public string CalculationId { get; set; }
        [DataMember] public string Date { get; set; }
        [DataMember] public string User { get; set; }
        [DataMember] public string Status { get; set; }
        [DataMember] public string Method { get; set; }
        [DataMember] public TepSettings Settings { get; set; }
        [DataMember] public List<TepResultRow> Results { get; set; } = new List<TepResultRow>();
        [DataMember] public List<TepDetailRow> Details { get; set; } = new List<TepDetailRow>();
        [DataMember] public List<TepWarningRow> Warnings { get; set; } = new List<TepWarningRow>();
        [DataMember] public List<long> CreatedElementIds { get; set; } = new List<long>();
        [DataMember] public List<TepCorrectionRow> Corrections { get; set; } = new List<TepCorrectionRow>();
    }

    [DataContract]
    internal sealed class TepCorrectionRow
    {
        [DataMember] public string Date { get; set; }
        [DataMember] public string Author { get; set; }
        [DataMember] public string Indicator { get; set; }
        [DataMember] public string Reason { get; set; }
        [DataMember] public double OldValue { get; set; }
        [DataMember] public double NewValue { get; set; }
    }

    public partial class AR_calculateTEP : Window
    {
        private readonly UIDocument _uidoc;
        private readonly ObservableCollection<TepIndicatorOption> _indicators = new ObservableCollection<TepIndicatorOption>();
        private readonly ObservableCollection<TepSourceOption> _sources = new ObservableCollection<TepSourceOption>();
        private readonly ObservableCollection<TepFormulaRule> _newRules = new ObservableCollection<TepFormulaRule>();
        private readonly ObservableCollection<TepFormulaRule> _oldRules = new ObservableCollection<TepFormulaRule>();
        private TepCalculationOutput _output;

        internal TepDialogAction RequestedAction { get; private set; }
        internal TepSettings SelectedSettings { get; private set; }
        internal string RequestedPath { get; private set; }

        internal AR_calculateTEP(UIDocument uidoc, IEnumerable<TepSourceOption> sources, TepSettings settings)
        {
            InitializeComponent();
            _uidoc = uidoc;
            foreach (TepSourceOption source in sources) _sources.Add(source);
            SourcesGrid.ItemsSource = _sources;
            LoadIndicators(settings);
            LoadFormulaRules(settings ?? new TepSettings());
            ApplySettings(settings ?? new TepSettings());
        }

        internal AR_calculateTEP(UIDocument uidoc, TepCalculationOutput output)
        {
            InitializeComponent();
            _uidoc = uidoc;
            _output = output;
            ConfigureResultsMode(output);
        }

        private void LoadIndicators(TepSettings settings)
        {
            string[,] data =
            {
                {"1", "Суммарная поэтажная площадь всего", "м²"}, {"2", "Суммарная поэтажная площадь жилых зданий", "м²"},
                {"2а", "Жилая часть суммарной поэтажной площади жилых зданий", "м²"}, {"2б", "Нежилая часть суммарной поэтажной площади жилых зданий", "м²"},
                {"3", "Суммарная поэтажная площадь нежилых зданий", "м²"}, {"4", "Площадь застройки", "м²"},
                {"5", "Строительный объём", "м³"}, {"5а", "Строительный объём выше отметки 0.000", "м³"}, {"5б", "Строительный объём ниже отметки 0.000", "м³"},
                {"6", "Общая площадь объекта", "м²"}, {"6а", "Общая площадь наземной части", "м²"}, {"6б", "Общая площадь подземной части", "м²"},
                {"7", "ННП встроенно-пристроенная", "м²"}, {"8", "ННП встроенно-пристроенная отдельно стоящая", "м²"},
                {"9", "Наземная площадь", "м²"}, {"9а", "Наземная площадь жилых зданий", "м²"}, {"9б", "Наземная площадь нежилых зданий", "м²"},
                {"10", "Расчётная площадь общественного здания", "м²"}, {"11", "Общая площадь квартир с летними помещениями", "м²"},
                {"12", "Площадь квартир без летних помещений", "м²"}, {"13", "Площадь помещений общественного назначения", "м²"},
                {"14", "Количество квартир", "шт."}, {"15", "Количество машино-мест в подземном паркинге", "шт."},
                {"16", "Этажность", "этажей"}, {"17", "Количество этажей", "этажей"}
            };
            bool hasSelection = settings != null && settings.SelectedIndicators != null && settings.SelectedIndicators.Count > 0;
            for (int i = 0; i < data.GetLength(0); ++i)
                _indicators.Add(new TepIndicatorOption { Number = data[i, 0], Name = data[i, 1], Unit = data[i, 2], IsSelected = !hasSelection || settings.SelectedIndicators.Contains(data[i, 0]) });
            IndicatorsGrid.ItemsSource = _indicators;
            AllIndicatorsCheckBox.IsChecked = _indicators.All(x => x.IsSelected);
        }

        private void ApplySettings(TepSettings s)
        {
            NewMethodRadio.IsChecked = !string.Equals(s.Method, "Old", StringComparison.OrdinalIgnoreCase);
            OldMethodRadio.IsChecked = !NewMethodRadio.IsChecked;
            RulesVersionTextBox.Text = s.RulesVersion;
            SelectTag(BuildingTypeCombo, s.BuildingType); SelectTag(AreaBasisCombo, s.AreaBasis); SelectTag(GroupingCombo, s.Grouping);
            SelectedBuildingsTextBox.Text = s.SelectedBuildings;
            ZeroElevationTextBox.Text = s.ZeroElevationMeters.ToString(CultureInfo.CurrentCulture); GroundElevationTextBox.Text = s.GroundElevationMeters.ToString(CultureInfo.CurrentCulture);
            ExcludedLevelsTextBox.Text = s.ExcludedLevelWords; BuildingParamTextBox.Text = s.BuildingParameter; BuildingTypeParamTextBox.Text = s.BuildingTypeParameter;
            FunctionParamTextBox.Text = s.FunctionParameter; ApartmentParamTextBox.Text = s.ApartmentParameter; SummerParamTextBox.Text = s.SummerParameter;
            SummerCoefParamTextBox.Text = s.SummerCoefficientParameter; ParkingParamTextBox.Text = s.ParkingParameter; BuiltInParamTextBox.Text = s.BuiltInParameter;
            StandaloneParamTextBox.Text = s.StandaloneParameter; IncludeParamTextBox.Text = s.IncludeParameter; ResidentialWordsTextBox.Text = s.ResidentialWords;
            PublicWordsTextBox.Text = s.PublicWords; TechnicalWordsTextBox.Text = s.TechnicalWords; SummerWordsTextBox.Text = s.SummerWords;
            ExcludedPublicWordsTextBox.Text = s.ExcludedPublicWords; UndergroundWordsTextBox.Text = s.UndergroundWords; FalseWordsTextBox.Text = s.FalseWords;
            RoundingCombo.SelectedIndex = Math.Max(0, Math.Min(3, s.RoundingDigits)); DetailedWarningsCheckBox.IsChecked = s.DetailedWarnings;
            ShowResultsCheckBox.IsChecked = s.ShowResults; CreateScheduleCheckBox.IsChecked = s.CreateSchedule; ExportCsvCheckBox.IsChecked = s.ExportCsv;
            ExportXlsxCheckBox.IsChecked = s.ExportXlsx; CreateViewsCheckBox.IsChecked = s.CreateViews; Create3DCheckBox.IsChecked = s.Create3D;
            UpdateGraphicsCheckBox.IsChecked = s.UpdateGraphics; PrefixTextBox.Text = s.Prefix;
            ManualCorrectionsTextBox.Text = s.ManualCorrections;
            if (s.IncludedSourceKeys != null && s.IncludedSourceKeys.Count > 0)
                foreach (TepSourceOption source in _sources) source.IsIncluded = s.IncludedSourceKeys.Contains(source.Key);
            if (s.GraphicsOnlySourceKeys != null)
                foreach (TepSourceOption source in _sources) source.GraphicsOnly = s.GraphicsOnlySourceKeys.Contains(source.Key);
            SourcesGrid.Items.Refresh();
            ShowSelectedRules();
        }

        private void LoadFormulaRules(TepSettings settings)
        {
            FillRules(_newRules, settings.NewMethodRules, false);
            FillRules(_oldRules, settings.OldMethodRules, true);
        }

        private static void FillRules(ObservableCollection<TepFormulaRule> target, IEnumerable<TepFormulaRule> stored, bool oldMethod)
        {
            target.Clear();
            List<TepFormulaRule> source = stored == null ? new List<TepFormulaRule>() : stored.Where(x => x != null).ToList();
            foreach (TepFormulaRule rule in DefaultFormulaRules(oldMethod))
            {
                TepFormulaRule saved = source.FirstOrDefault(x => string.Equals(x.Number, rule.Number, StringComparison.OrdinalIgnoreCase));
                target.Add(new TepFormulaRule { Number = rule.Number, Name = rule.Name, Expression = saved == null || string.IsNullOrWhiteSpace(saved.Expression) ? rule.Expression : saved.Expression });
            }
        }

        private static IEnumerable<TepFormulaRule> DefaultFormulaRules(bool oldMethod)
        {
            string[,] data =
            {
                {"1","Суммарная поэтажная площадь всего","A_ALL"},{"2","Суммарная поэтажная площадь жилых зданий","A_RES_BUILDING"},{"2а","Жилая часть площади жилых зданий","A_RES_IN_RES_BUILDING"},{"2б","Нежилая часть площади жилых зданий","A_NONRES_IN_RES_BUILDING"},{"3","Площадь нежилых зданий","A_NONRES_BUILDING"},
                {"4","Площадь застройки","FOOTPRINT"},{"5","Строительный объём","V_ALL"},{"5а","Объём выше 0.000","V_ABOVE"},{"5б","Объём ниже 0.000","V_BELOW"},{"6","Общая площадь","A_ALL"},{"6а","Наземная площадь","A_ABOVE"},{"6б","Подземная площадь","A_BELOW"},
                {"7","ННП встроенно-пристроенная","A_NNP_BUILTIN"},{"8","ННП отдельно стоящая","A_NNP_STANDALONE"},{"9","Наземная площадь","A_ABOVE"},{"9а","Наземная площадь жилых зданий","A_ABOVE_RES_BUILDING"},{"9б","Наземная площадь нежилых зданий","A_ABOVE_NONRES_BUILDING"},{"10","Расчётная общественная площадь","A_PUBLIC_CALC"},
                {"11","Площадь квартир с коэффициентами","A_APT_COEF"},{"12","Площадь квартир без летних помещений","A_APT_HEATED"},{"13","Общественные помещения","A_PUBLIC"},{"14","Количество квартир","COUNT_APT"},{"15","Количество машино-мест","COUNT_PARKING"},{"16","Этажность","FLOORS_ABOVE"},{"17","Количество этажей","FLOORS_ALL"}
            };
            for (int i = 0; i < data.GetLength(0); ++i) yield return new TepFormulaRule { Number = data[i, 0], Name = data[i, 1], Expression = data[i, 2] };
        }

        private void ShowSelectedRules()
        {
            if (FormulaRulesGrid == null) return;
            FormulaRulesGrid.ItemsSource = OldMethodRadio.IsChecked == true ? _oldRules : _newRules;
        }

        private TepSettings ReadSettings()
        {
            FormulaRulesGrid.CommitEdit(DataGridEditingUnit.Cell, true);
            double zero, ground;
            if (!TryNumber(ZeroElevationTextBox.Text, out zero) || !TryNumber(GroundElevationTextBox.Text, out ground))
                throw new InvalidOperationException("Отметки 0.000 и земли должны быть числами в метрах.");
            return new TepSettings
            {
                Method = OldMethodRadio.IsChecked == true ? "Old" : "New",
                RulesVersion = RulesVersionTextBox.Text.Trim(),
                BuildingType = SelectedTag(BuildingTypeCombo),
                AreaBasis = SelectedTag(AreaBasisCombo),
                Grouping = SelectedTag(GroupingCombo),
                SelectedBuildings = SelectedBuildingsTextBox.Text,
                ZeroElevationMeters = zero,
                GroundElevationMeters = ground,
                ExcludedLevelWords = ExcludedLevelsTextBox.Text,
                BuildingParameter = BuildingParamTextBox.Text.Trim(),
                BuildingTypeParameter = BuildingTypeParamTextBox.Text.Trim(),
                FunctionParameter = FunctionParamTextBox.Text.Trim(),
                ApartmentParameter = ApartmentParamTextBox.Text.Trim(),
                SummerParameter = SummerParamTextBox.Text.Trim(),
                SummerCoefficientParameter = SummerCoefParamTextBox.Text.Trim(),
                ParkingParameter = ParkingParamTextBox.Text.Trim(),
                BuiltInParameter = BuiltInParamTextBox.Text.Trim(),
                StandaloneParameter = StandaloneParamTextBox.Text.Trim(),
                IncludeParameter = IncludeParamTextBox.Text.Trim(),
                ResidentialWords = ResidentialWordsTextBox.Text,
                PublicWords = PublicWordsTextBox.Text,
                TechnicalWords = TechnicalWordsTextBox.Text,
                SummerWords = SummerWordsTextBox.Text,
                ExcludedPublicWords = ExcludedPublicWordsTextBox.Text,
                UndergroundWords = UndergroundWordsTextBox.Text,
                FalseWords = FalseWordsTextBox.Text,
                RoundingDigits = Math.Max(0, RoundingCombo.SelectedIndex),
                DetailedWarnings = DetailedWarningsCheckBox.IsChecked == true,
                ShowResults = ShowResultsCheckBox.IsChecked == true,
                CreateSchedule = CreateScheduleCheckBox.IsChecked == true,
                ExportCsv = ExportCsvCheckBox.IsChecked == true,
                ExportXlsx = ExportXlsxCheckBox.IsChecked == true,
                CreateViews = CreateViewsCheckBox.IsChecked == true,
                Create3D = Create3DCheckBox.IsChecked == true,
                UpdateGraphics = UpdateGraphicsCheckBox.IsChecked == true,
                Prefix = string.IsNullOrWhiteSpace(PrefixTextBox.Text) ? "ТЭП_" : PrefixTextBox.Text.Trim(),
                ManualCorrections = ManualCorrectionsTextBox.Text,
                SelectedIndicators = _indicators.Where(x => x.IsSelected).Select(x => x.Number).ToList(),
                IncludedSourceKeys = _sources.Where(x => x.IsIncluded).Select(x => x.Key).ToList(),
                GraphicsOnlySourceKeys = _sources.Where(x => x.GraphicsOnly).Select(x => x.Key).ToList(),
                NewMethodRules = _newRules.Select(CloneRule).ToList(),
                OldMethodRules = _oldRules.Select(CloneRule).ToList()
            };
        }

        private static TepFormulaRule CloneRule(TepFormulaRule rule) { return new TepFormulaRule { Number = rule.Number, Name = rule.Name, Expression = rule.Expression }; }

        private void ConfigureResultsMode(TepCalculationOutput output)
        {
            IndicatorsTab.Visibility = System.Windows.Visibility.Collapsed;
            foreach (TabItem tab in MainTabs.Items.Cast<TabItem>().Where(x => x != ResultsTab && x != DetailsTab && x != WarningsTab)) tab.Visibility = System.Windows.Visibility.Collapsed;
            ResultsTab.Visibility = DetailsTab.Visibility = WarningsTab.Visibility = System.Windows.Visibility.Visible;
            ResultsGrid.ItemsSource = output.Results; DetailsGrid.ItemsSource = output.Details; WarningsGrid.ItemsSource = output.Warnings;
            ResultStatusText.Text = string.Format("{0}. Расчёт {1}; {2}; строк результатов: {3}; предупреждений: {4}.", output.Status, output.CalculationId, output.Date, output.Results.Count, output.Warnings.Count);
            CalculateButton.Visibility = LastResultButton.Visibility = ImportButton.Visibility = ExportSettingsButton.Visibility = ClearButton.Visibility = System.Windows.Visibility.Collapsed;
            ExportCsvButton.Visibility = ExportXlsxButton.Visibility = System.Windows.Visibility.Visible; MainTabs.SelectedItem = ResultsTab;
        }

        private void CalculateButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                IndicatorsGrid.CommitEdit(DataGridEditingUnit.Cell, true); SourcesGrid.CommitEdit(DataGridEditingUnit.Cell, true);
                SelectedSettings = ReadSettings();
                if (SelectedSettings.SelectedIndicators.Count == 0) throw new InvalidOperationException("Выберите хотя бы один показатель.");
                if (SelectedSettings.IncludedSourceKeys.Count == 0) throw new InvalidOperationException("Выберите хотя бы один доступный источник.");
                RequestedAction = TepDialogAction.Calculate; DialogResult = true;
            }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, "Расчёт ТЭП", MessageBoxButton.OK, MessageBoxImage.Warning); }
        }

        private void AllIndicatorsCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            if (IndicatorsGrid == null || _indicators.Count == 0) return;
            bool value = AllIndicatorsCheckBox.IsChecked == true;
            foreach (TepIndicatorOption row in _indicators) row.IsSelected = value;
            IndicatorsGrid.Items.Refresh();
        }

        private void MethodRadio_Checked(object sender, RoutedEventArgs e) { ShowSelectedRules(); }
        private void ResetFormulaRulesButton_Click(object sender, RoutedEventArgs e)
        {
            bool oldMethod = OldMethodRadio.IsChecked == true;
            FillRules(oldMethod ? _oldRules : _newRules, null, oldMethod);
            ShowSelectedRules();
        }

        private void LastResultButton_Click(object sender, RoutedEventArgs e) { RequestedAction = TepDialogAction.ShowLast; DialogResult = true; }
        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show(this, "Удалить только объекты и сохранённые результаты, созданные командой ТЭП?", "Очистка ТЭП", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes) return;
            RequestedAction = TepDialogAction.Clear; DialogResult = true;
        }
        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog d = new OpenFileDialog { Filter = "JSON (*.json)|*.json", Title = "Импорт настроек ТЭП" };
            if (d.ShowDialog(this) != true) return; RequestedPath = d.FileName; RequestedAction = TepDialogAction.ImportSettings; DialogResult = true;
        }
        private void ExportSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            try { SelectedSettings = ReadSettings(); } catch (Exception ex) { MessageBox.Show(this, ex.Message); return; }
            SaveFileDialog d = new SaveFileDialog { Filter = "JSON (*.json)|*.json", FileName = "ТЭП_настройки.json", Title = "Экспорт настроек ТЭП" };
            if (d.ShowDialog(this) != true) return; RequestedPath = d.FileName; RequestedAction = TepDialogAction.ExportSettings; DialogResult = true;
        }
        private void ExportCsvButton_Click(object sender, RoutedEventArgs e) { RequestResultExport(TepDialogAction.ExportCsv, "CSV (*.csv)|*.csv", "ТЭП_результаты.csv"); }
        private void ExportXlsxButton_Click(object sender, RoutedEventArgs e) { RequestResultExport(TepDialogAction.ExportXlsx, "Excel (*.xlsx)|*.xlsx", "ТЭП_результаты.xlsx"); }
        private void RequestResultExport(TepDialogAction action, string filter, string fileName)
        {
            SaveFileDialog d = new SaveFileDialog { Filter = filter, FileName = fileName };
            if (d.ShowDialog(this) != true) return; RequestedPath = d.FileName; RequestedAction = action; DialogResult = true;
        }
        private void CloseButton_Click(object sender, RoutedEventArgs e) { RequestedAction = TepDialogAction.None; Close(); }

        private void DetailsGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e) { Navigate(DetailsGrid.SelectedItem as TepDetailRow); }
        private void WarningsGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e) { Navigate(WarningsGrid.SelectedItem as TepWarningRow); }
        private void Navigate(TepDetailRow row) { if (row != null) Navigate(row.ElementId, row.HostLinkId); }
        private void Navigate(TepWarningRow row) { if (row != null) Navigate(row.ElementId, row.HostLinkId); }
        private void Navigate(long elementId, long hostLinkId)
        {
            try
            {
                long id = hostLinkId > 0 ? hostLinkId : elementId; if (id <= 0) return;
                ElementId revitId = IDHelper.CreateElementId(id); _uidoc.Selection.SetElementIds(new[] { revitId }); _uidoc.ShowElements(revitId);
            }
            catch (Exception ex) { MessageBox.Show(this, "Не удалось перейти к объекту: " + ex.Message, "Расчёт ТЭП", MessageBoxButton.OK, MessageBoxImage.Information); }
        }

        private static string SelectedTag(WpfComboBox box) { ComboBoxItem item = box.SelectedItem as ComboBoxItem; return item == null ? string.Empty : Convert.ToString(item.Tag, CultureInfo.InvariantCulture); }
        private static void SelectTag(WpfComboBox box, string tag)
        {
            foreach (ComboBoxItem item in box.Items) if (string.Equals(Convert.ToString(item.Tag), tag, StringComparison.OrdinalIgnoreCase)) { box.SelectedItem = item; return; }
            box.SelectedIndex = 0;
        }
        private static bool TryNumber(string text, out double value) { return double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out value) || double.TryParse(text.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out value); }
    }
}