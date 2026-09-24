using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using Microsoft.Win32;
using TEP = KPLN_Tools.ExternalCommands.Command_AR_CalculateTEP;

namespace KPLN_Tools.Forms
{
    public partial class AR_CalculateTEP : Window
    {
        public TEP.Engine Engine { get; private set; }
        public TEP.Settings Config { get { return Engine.Config; } }
        public List<TEP.Choice> Methods { get { return TEP.Choices("method"); } }
        public List<TEP.Choice> Profiles { get { return TEP.Choices("profile"); } }
        public List<TEP.Choice> GroupModes { get { return TEP.Choices("group"); } }
        public List<TEP.Choice> DatumModes { get { return TEP.Choices("datum"); } }
        public List<TEP.Choice> ZeroModes { get { return DatumModes.Where(x => x.Key != "terrain").ToList(); } }
        public List<TEP.Choice> GeometryModes { get { return TEP.Choices("geometry"); } }
        public List<TEP.Choice> CountModes { get { return TEP.Choices("count"); } }
        public List<TEP.Choice> GraphicsModes { get { return TEP.Choices("graphics").Where(x => x.Key != "none").ToList(); } }
        public List<TEP.Choice> DiagnosticModes { get { return TEP.Choices("diagnostics"); } }
        public List<TEP.Choice> MetricChoices { get { return new List<TEP.Choice> { new TEP.Choice("all", "Все показатели", "Общая классификация применяется до правил конкретного показателя.") }.Concat(TEP.Catalog().Select(m => new TEP.Choice(m.Key, m.Name, m.Description))).ToList(); } }
        public List<TEP.Metric> ContourMetrics { get { return Config.Metrics.Where(m => TEP.SupportsContours(m.Key)).ToList(); } }
        public List<TEP.Choice> ContourModes { get { return TEP.Choices("contour"); } }
        public List<TEP.Choice> WallSelectionModes { get { return TEP.Choices("wall-selection"); } }
        public List<TEP.Choice> WallBoundaries { get { return TEP.Choices("wall-boundary"); } }
        public List<TEP.Choice> ContourExclusionModes { get { return TEP.Choices("contour-exclusions"); } }
        public List<TEP.Choice> MaskHeightModes { get { return TEP.Choices("mask-height"); } }
        private static readonly Brush ErrorTextBrush = new SolidColorBrush(Color.FromRgb(180, 35, 24));
        private bool changing, busy, cancelRequested;
        private readonly System.Diagnostics.Stopwatch progressClock = System.Diagnostics.Stopwatch.StartNew();
        private List<TEP.Issue> displayedIssues;
        public AR_CalculateTEP(TEP.Engine engine)
        {
            Engine = engine; InitializeComponent(); DataContext = this; ConfigureTables(); RestoreContourSelection();
            Loaded += (s, e) => Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(() =>
            {
                if (Engine.IsInitialized) { ShowReport(Engine.Last); UpdateStep(); if (Engine.ReopenContours) { Steps.SelectedIndex = 0; CalculationTabs.SelectedIndex = 1; Engine.ReopenContours = false; } return; }
                SetBusy(true);
                try { Engine.Initialize(ReportProgress); Rebind(); ConfigureTables(); ShowReport(Engine.Last); Status.Text = Engine.Last == null ? "Модель загружена. Настройте состав расчёта." : "Загружен сохранённый отчёт. После изменения модели выполните новый расчёт."; }
                catch (System.OperationCanceledException) { SetBusy(false); Close(); return; }
                catch (Exception ex) { SetBusy(false); MessageBox.Show(this, ex.Message, "Не удалось загрузить модель ТЭП", MessageBoxButton.OK, MessageBoxImage.Error); Close(); return; }
                finally { SetBusy(false); }
            }));
            Closing += (s, e) => { if (busy) e.Cancel = true; };
        }
        private void SetBusy(bool value)
        {
            busy = value; cancelRequested = false; Steps.IsEnabled = !value; HeaderActions.IsEnabled = !value;
            CalculateButton.IsEnabled = !value; BackButton.IsEnabled = !value; NextButton.IsEnabled = !value;
            CreateViewsCheckBox.IsEnabled = !value;
            CancelButton.Visibility = value ? Visibility.Visible : Visibility.Collapsed; CancelButton.IsEnabled = value;
            WorkProgress.Visibility = value ? Visibility.Visible : Visibility.Collapsed;
            if (!value) UpdateStep(); else progressClock.Restart();
        }
        private void Cancel_Click(object sender, RoutedEventArgs e)
        { cancelRequested = true; CancelButton.IsEnabled = false; Status.Text = "Отмена после текущей операции Revit..."; }
        private void ReportProgress(string message)
        {
            if (cancelRequested) throw new System.OperationCanceledException();
            if (progressClock.ElapsedMilliseconds < 80) return;
            Status.Text = message; progressClock.Restart();
            // Revit API remains on the command thread. Pump only while all editing actions are disabled.
            var frame = new DispatcherFrame();
            Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(() => frame.Continue = false));
            Dispatcher.PushFrame(frame);
            if (cancelRequested) throw new System.OperationCanceledException();
        }
        private void Help_Click(object sender, RoutedEventArgs e)
        {
            var text = new TextBlock { Text = Help, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(24), FontSize = 14 };
            new Window
            {
                Title = "ТЭП | Порядок работы",
                Owner = this,
                Width = 760,
                Height = 650,
                MinWidth = 480,
                MinHeight = 360,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Content = new ScrollViewer { Content = text, VerticalScrollBarVisibility = ScrollBarVisibility.Auto }
            }.ShowDialog();
        }
        private static Binding Bind(string path, string format = null)
        {
            var binding = new Binding(path) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged, StringFormat = format, ValidatesOnExceptions = true, NotifyOnValidationError = true };
            int digits; if (format != null && format.StartsWith("{0:N") && int.TryParse(format.Substring(4).TrimEnd('}'), out digits)) binding.Converter = new RoundedNumber { Digits = digits };
            return binding;
        }
        private sealed class RoundedNumber : IValueConverter
        {
            internal int Digits;
            public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
            { return value is double ? Math.Round((double)value, Digits, MidpointRounding.AwayFromZero) : value; }
            public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) { return Binding.DoNothing; }
        }
        private void RefreshPrecision(DataGrid grid)
        { foreach (var column in grid.Columns.OfType<DataGridTextColumn>()) if ((column.Binding as Binding)?.Path.Path == "Value") { var binding = Bind("Value", "{0:N" + Config.Decimals + "}"); binding.Mode = BindingMode.OneWay; column.Binding = binding; } }
        private static string ColumnTip(DataGrid grid, string title, string path, bool editable, string extra = null)
        {
            if (grid?.Name == "ContoursGrid")
            {
                switch (path)
                {
                    case "ElementLabel": return "ID цветовой области активной модели. Расчёт читает её актуальные границы при каждом запуске.";
                    case "Kind": return "Основной контур задаёт исходную площадь в ручном режиме; вырез вычитается также в режиме наружных стен.";
                    case "LevelName": return "Уровень плана, на котором находится область. Изменение уровня требует повторного закрепления области.";
                    case "Building": return "Имя расчётного корпуса. Вырезы помещений применяются при совпадении корпуса и секции с основным контуром.";
                    case "Section": return "Имя секции. Пустое значение означает секцию без имени; оно не является подстановкой для всех секций.";
                    case "Profile": return "Профиль здания для проверки правил включения площади этажа и нормативных исключений.";
                    case "BuildingClass": return "Жилое или нежилое здание. Определяет принадлежность поэтажных площадей к соответствующему показателю.";
                    case "Height": return "Высота вертикального участка в метрах. Пусто - до верха участка, заданного таблицей этажей или следующим включённым уровнем, с учётом смещения низа. Для последнего участка высота обязательна; для плоских площадей не применяется.";
                    case "BottomOffset": return "Смещение низа ручного выреза относительно его уровня, в метрах. Положительное поднимает, отрицательное опускает вырез. У основного контура должно быть нулём; для площадей не применяется.";
                }
                return extra ?? title;
            }
            string meaning = null, example = null;
            switch (path)
            {
                case "Source": meaning = grid == null ? "" : grid.Name == "BuildingsGrid" ? "Укажите имя источника или *. Ограничивает строку соответствия выбранной моделью; * действует для всех источников." : "Укажите имя модели или экземпляра связи из списка. Определяет, к какому источнику относится корректировка."; example = "Корпус 1.rvt"; break;
                case "MatchValue": meaning = "Укажите исходное значение корпуса, рабочего набора или имени связи либо *. Совпадение назначает объектам корпус и профиль этой строки."; example = "Секция А"; break;
                case "Building": meaning = grid.Name == "LevelsGrid" ? "Укажите корпус уточнения уровня. Пусто - общая настройка; имя ограничивает уточнение одним корпусом." : "Введите итоговое имя корпуса. Под этим именем объединяются площади и другие показатели здания."; example = "Корпус 1"; break;
                case "Section": meaning = "Укажите секцию уточнения уровня. Позволяет отдельно задать наземность и параметры этажа этой секции; пусто - все секции."; example = "А"; break;
                case "TopSlab": meaning = "Введите абсолютную отметку верха перекрытия в метрах основной модели. Используется при определении наземности цоколя."; example = "158,20"; break;
                case "Height": meaning = "Введите высоту этажа в метрах. Используется в правилах включения технического этажа / надстройки; пустое значение требует данных модели."; example = "2,40"; break;
                case "RoofRatio": meaning = "Введите долю площади надстройки от кровли от 0 до 1. Влияет на включение технической надстройки общественного здания."; example = "0,15"; break;
                case "RoofArea": meaning = "Введите площадь надстройки последнего верхнего этажа в м². Вместе с высотой применяется для правил высотного здания."; example = "7,5"; break;
                case "Name": meaning = "Введите точное имя параметра экземпляра или типа для назначения, указанного в этой строке. Правило чтения и влияние описаны в соседнем столбце; пусто - параметр не назначен."; example = "ТЭП_Назначение"; break;
                case "Category": meaning = "Выберите или введите точное имя категории. Правило проверяет только эту категорию; пусто - все категории."; example = "Помещения"; break;
                case "Parameter": meaning = "Выберите или введите имя параметра экземпляра / типа. По его значению правило определяет назначение и включение объекта; @Name означает имя объекта."; example = "ТЭП_Назначение"; break;
                case "Value": meaning = grid.Name == "CorrectionsGrid" ? "Введите значение для выбранного действия: назначение, имя корпуса или числовую дельту в единицах показателя. Корректирует только расчёт." : "Введите сравниваемое значение параметра. Используется условием «Равно» или «Содержит»; для «Заполнен» / «Пусто» не требуется."; example = grid.Name == "CorrectionsGrid" ? "12,5 (дельта площади в м²)" : "Лоджия"; break;
                case "Coefficient": meaning = "Введите коэффициент от 0 до 1 либо оставьте пустым для нормативного значения назначения. Меняет площадь квартиры с летними помещениями; параметр коэффициента объекта имеет приоритет."; example = "0,5"; break;
                case "Element": meaning = "Введите ElementId или UniqueId объекта. Для числовой дельты укажите корпус, для дельты этажности - Корпус|Секция. Определяет объект корректировки."; example = "123456"; break;
                case "Reason": meaning = "Введите причину ручного решения. Обоснование сохраняется в отчёте вместе с автором, датой и прежним значением."; example = "Исключён дублирующий контур лоджии"; break;
                case "AreaScheme": meaning = "Выберите или введите схему зон этого показателя. Пусто - общая схема; отдельное значение позволяет считать разные показатели по разным контурам."; example = "ГНС - наружный контур"; break;
                case "Mode": meaning = "Выберите участие источника. Включение даёт вклад в итог; исключение убирает источник; сверка оставляет только проверочную графику."; break;
                case "Profile": meaning = "Выберите нормативный профиль источника или корпуса. Определяет включение помещений и этажей при соответствующем способе выбора типа здания."; break;
                case "Class": meaning = "Выберите класс здания. От него зависит отнесение ГНС и НП к жилым или нежилым зданиям."; break;
                case "Kind": meaning = "Выберите вид этажа. Меняет нормативные правила включения его площадей и этажности."; break;
                case "Above": meaning = "Выберите наземность этажа. Автоматический режим использует землю и параметры этажа; ручной вариант переопределяет его для этого уровня."; break;
                case "Geometry": meaning = "Выберите источник площади показателя. Пустой выбор наследует общую настройку; Rooms, Spaces и Areas не должны дублировать один этаж."; break;
                case "Metric": meaning = "Выберите показатель, на который действует строка. «Все показатели» задаёт общую классификацию / корректировку до правил отдельного показателя."; break;
                case "Operation": meaning = "Выберите условие сравнения параметра. При совпадении применяется действие и назначение этой строки."; break;
                case "Action": meaning = grid.Name == "CorrectionsGrid" ? "Выберите вид ручной корректировки. Действие определяет смысл полей объекта и нового значения." : "Выберите действие правила. Классификация задаёт назначение; включение или исключение переопределяет участие объекта."; break;
                case "Role": meaning = "Выберите назначение совпавших объектов. От него зависят нормативное включение и коэффициент площади."; break;
                case "Part": meaning = "Выберите часть здания. Определяет распределение площади жилого здания на жилую и нежилую части."; break;
                case "Loaded": meaning = "Показывает, доступна ли модель связи. Незагруженный источник нельзя рассчитать; загрузите связь в Revit либо исключите её."; break;
                case "Enabled": meaning = "Включите флажок, чтобы правило участвовало в классификации объектов. Снятый флажок отключает строку без удаления."; break;
                case "Include": meaning = grid.Name == "LevelsGrid" ? "Включите флажок, чтобы уровень участвовал в расчёте. Снятие исключает его из состава расчётных этажей." : "Включите флажок, чтобы учитывать объекты этого соответствия корпуса. Снятие исключает совпавшие объекты."; break;
            }
            if (meaning == null) meaning = extra ?? ("«" + title + "»: значение в строке результата.");
            if (!string.IsNullOrWhiteSpace(extra) && !meaning.Contains(extra)) meaning += "\n" + extra;
            return meaning + (editable && example != null ? "\nПример: " + example : "");
        }
        private static void TextColumn(DataGrid grid, string title, string path, double width = 150, bool readOnly = false, string format = null, string tip = null)
        {
            tip = readOnly ? (tip ?? ("«" + title + "»: значение в строке результата.")) : ColumnTip(grid, title, path, true, tip);
            var header = new TextBlock { Text = title, ToolTip = tip, VerticalAlignment = VerticalAlignment.Center };
            var binding = Bind(path, format); if (readOnly) binding.Mode = BindingMode.OneWay;
            var elementStyle = new Style(typeof(TextBlock)); elementStyle.Setters.Add(new Setter(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center));
            elementStyle.Setters.Add(new Setter(FrameworkElement.MarginProperty, new Thickness(7, 0, 7, 0))); elementStyle.Setters.Add(new Setter(TextBlock.TextTrimmingProperty, TextTrimming.CharacterEllipsis));
            elementStyle.Setters.Add(new Setter(FrameworkElement.ToolTipProperty, readOnly ? (object)new Binding(path) : tip));
            string errorPath = grid.Name == "IssuesGrid" ? "Severity" : grid.Name == "SummaryGrid" ? "Status" : grid.Name == "DetailsGrid" ? "HasError" : null;
            if (errorPath != null)
            {
                var error = new DataTrigger { Binding = new Binding(errorPath), Value = errorPath == "HasError" ? (object)true : errorPath == "Severity" ? "Ошибка" : "Неполный результат" };
                error.Setters.Add(new Setter(TextBlock.ForegroundProperty, ErrorTextBrush)); elementStyle.Triggers.Add(error);
            }
            var editStyle = new Style(typeof(TextBox)); editStyle.Setters.Add(new Setter(FrameworkElement.ToolTipProperty, tip)); editStyle.Setters.Add(new Setter(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center)); editStyle.Setters.Add(new Setter(FrameworkElement.MinHeightProperty, 28.0)); editStyle.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(4, 2, 4, 2)));
            grid.Columns.Add(new DataGridTextColumn { Header = header, Binding = binding, Width = width, IsReadOnly = readOnly, ElementStyle = elementStyle, EditingElementStyle = editStyle });
        }
        private static void CheckColumn(DataGrid grid, string title, string path, double width = 70, bool readOnly = false)
        {
            var binding = Bind(path); if (readOnly) binding.Mode = BindingMode.OneWay;
            var style = new Style(typeof(CheckBox)); style.Setters.Add(new Setter(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center));
            style.Setters.Add(new Setter(FrameworkElement.HorizontalAlignmentProperty, HorizontalAlignment.Center)); style.Setters.Add(new Setter(FrameworkElement.MarginProperty, new Thickness(0)));
            style.Setters.Add(new Setter(FrameworkElement.ToolTipProperty, ColumnTip(grid, title, path, false)));
            var display = new Style(typeof(CheckBox), style); display.Setters.Add(new Setter(UIElement.IsHitTestVisibleProperty, false)); display.Setters.Add(new Setter(UIElement.FocusableProperty, false));
            grid.Columns.Add(new DataGridCheckBoxColumn { Header = new TextBlock { Text = title, ToolTip = ColumnTip(grid, title, path, false) }, Binding = binding, Width = width, IsReadOnly = readOnly, ElementStyle = display, EditingElementStyle = style });
        }
        private static void ChoiceColumn(DataGrid grid, string title, string path, IEnumerable<TEP.Choice> choices, double width = 180)
        {
            var combo = new FrameworkElementFactory(typeof(ComboBox)); combo.SetValue(ItemsControl.ItemsSourceProperty, choices.ToList());
            combo.SetValue(ItemsControl.DisplayMemberPathProperty, "Label"); combo.SetValue(System.Windows.Controls.Primitives.Selector.SelectedValuePathProperty, "Key");
            combo.SetBinding(System.Windows.Controls.Primitives.Selector.SelectedValueProperty, Bind(path));
            combo.SetValue(Control.PaddingProperty, new Thickness(4, 2, 4, 2)); combo.SetValue(FrameworkElement.MinHeightProperty, 28.0); combo.SetValue(FrameworkElement.HeightProperty, 28.0); combo.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center); combo.SetValue(FrameworkElement.MarginProperty, new Thickness(4, 0, 4, 0));
            combo.SetValue(FrameworkElement.TagProperty, ColumnTip(grid, title, path, false));
            var template = new DataTemplate { VisualTree = combo }; grid.Columns.Add(new DataGridTemplateColumn { Header = new TextBlock { Text = title, ToolTip = ColumnTip(grid, title, path, false) }, CellTemplate = template, Width = width });
        }
        private static void EditableColumn(DataGrid grid, string title, string path, IEnumerable<string> values, double width = 220, string tooltip = null)
        {
            var combo = new FrameworkElementFactory(typeof(ComboBox)); combo.SetValue(ItemsControl.ItemsSourceProperty, values.ToList()); combo.SetValue(ComboBox.IsEditableProperty, true);
            combo.SetBinding(ComboBox.TextProperty, Bind(path)); combo.SetValue(Control.PaddingProperty, new Thickness(4, 2, 4, 2)); combo.SetValue(FrameworkElement.MinHeightProperty, 28.0); combo.SetValue(FrameworkElement.HeightProperty, 28.0); combo.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center); combo.SetValue(FrameworkElement.MarginProperty, new Thickness(4, 0, 4, 0));
            tooltip = ColumnTip(grid, title, path, true, tooltip);
            if (grid.Name == "ParametersGrid")
            { var tip = new MultiBinding { StringFormat = "{0}\n{1}" }; tip.Bindings.Add(new Binding { Source = tooltip }); tip.Bindings.Add(new Binding("Description")); combo.SetBinding(FrameworkElement.ToolTipProperty, tip); }
            else combo.SetValue(FrameworkElement.ToolTipProperty, tooltip);
            grid.Columns.Add(new DataGridTemplateColumn { Header = new TextBlock { Text = title, ToolTip = tooltip }, CellTemplate = new DataTemplate { VisualTree = combo }, Width = width });
        }
        private void ConfigureTables()
        {
            foreach (var grid in new[] { ContoursGrid, MetricSourcesGrid, SourcesGrid, BuildingsGrid, LevelsGrid, ParametersGrid, RulesGrid, CorrectionsGrid, SummaryGrid, DetailsGrid, IssuesGrid }) grid.Columns.Clear();
            TextColumn(ContoursGrid, "Область ID", "ElementLabel", 115, true);
            ChoiceColumn(ContoursGrid, "Назначение", "Kind", TEP.Choices("sketch-kind"), 175);
            TextColumn(ContoursGrid, "Уровень", "LevelName", 150, true); TextColumn(ContoursGrid, "Корпус", "Building", 180); TextColumn(ContoursGrid, "Секция", "Section", 110);
            ChoiceColumn(ContoursGrid, "Профиль", "Profile", Profiles.Where(x => !x.Key.StartsWith("by-") && x.Key != "high-mixed"), 230);
            ChoiceColumn(ContoursGrid, "Класс здания", "BuildingClass", new[] { new TEP.Choice("residential", "Жилое", "Площадь относится к жилому зданию."), new TEP.Choice("nonresidential", "Нежилое", "Площадь относится к нежилому зданию.") }, 140);
            TextColumn(ContoursGrid, "Высота, м", "Height", 120); TextColumn(ContoursGrid, "Смещение низа, м", "BottomOffset", 155);
            TextColumn(MetricSourcesGrid, "Показатель", "Name", 510, true);
            ChoiceColumn(MetricSourcesGrid, "Источник площадей", "Geometry", new[] { new TEP.Choice("", "Общая настройка", "Использует источник площадей, выбранный на вкладке «Параметры».") }.Concat(GeometryModes), 235);
            EditableColumn(MetricSourcesGrid, "Схема зон", "AreaScheme", Engine.AreaSchemes, 270, "Отдельная схема именно этого показателя. Пусто - общая схема; при нескольких схемах требуется явный выбор.");
            TextColumn(SourcesGrid, "Источник / экземпляр связи", "Name", 400, true); CheckColumn(SourcesGrid, "Загружен", "Loaded", 90, true);
            ChoiceColumn(SourcesGrid, "Участие", "Mode", TEP.Choices("source"), 240); ChoiceColumn(SourcesGrid, "Профиль", "Profile", Profiles.Where(x => !x.Key.StartsWith("by-")), 260);
            EditableColumn(BuildingsGrid, "Источник или *", "Source", Engine.Sources.Select(s => s.Name).Concat(new[] { "*" }), 260);
            CheckColumn(BuildingsGrid, "Учесть", "Include");
            TextColumn(BuildingsGrid, "Исходное значение или *", "MatchValue", 190); TextColumn(BuildingsGrid, "Корпус", "Building", 180);
            ChoiceColumn(BuildingsGrid, "Профиль", "Profile", Profiles.Where(x => !x.Key.StartsWith("by-")), 240);
            ChoiceColumn(BuildingsGrid, "Класс здания", "Class", new[] { new TEP.Choice("residential", "Жилое", "ГНС и НП относятся к жилому зданию, даже если часть помещений нежилая."), new TEP.Choice("nonresidential", "Нежилое", "ГНС и НП относятся к нежилому зданию.") }, 150);
            CheckColumn(LevelsGrid, "Учесть", "Include"); TextColumn(LevelsGrid, "Источник", "Source", 210, true); TextColumn(LevelsGrid, "Уровень", "Name", 170, true); TextColumn(LevelsGrid, "Отметка, м", "ElevationMeters", 110, true, "{0:0.000}");
            TextColumn(LevelsGrid, "Корпус уточнения", "Building", 150); TextColumn(LevelsGrid, "Секция уточнения", "Section", 150);
            ChoiceColumn(LevelsGrid, "Вид этажа", "Kind", TEP.Choices("level"), 200); ChoiceColumn(LevelsGrid, "Наземность", "Above", TEP.Choices("above"), 170);
            TextColumn(LevelsGrid, "Верх перекрытия, м", "TopSlab", 150, tip: "Абсолютная отметка в координатах основной модели. Для жилого цоколя проверяется превышение над землёй на 2 м.");
            TextColumn(LevelsGrid, "Высота, м", "Height", 110); TextColumn(LevelsGrid, "Доля от кровли", "RoofRatio", 130, tip: "Число от 0 до 1. Используется для технической надстройки по правилам общественного здания.");
            TextColumn(LevelsGrid, "Площадь надстройки, м²", "RoofArea", 170, tip: "Суммарная площадь надстройки на последнем верхнем этаже высотного здания; используется совместно с высотой для порогов 8 м² и 2,5 м.");
            TextColumn(ParametersGrid, "Назначение", "Title", 360, true); EditableColumn(ParametersGrid, "Имя параметра экземпляра / типа", "Name", Engine.Parameters, 360, "Выберите существующий параметр либо введите точное имя. При нескольких параметрах с одним именем расчёт сообщит неоднозначность."); TextColumn(ParametersGrid, "Правило чтения", "Description", 700, true);
            CheckColumn(RulesGrid, "Вкл.", "Enabled", 55); ChoiceColumn(RulesGrid, "Показатель", "Metric", MetricChoices, 270); EditableColumn(RulesGrid, "Категория (пусто = все)", "Category", Engine.Categories, 190);
            EditableColumn(RulesGrid, "Параметр", "Parameter", Engine.Parameters, 220); ChoiceColumn(RulesGrid, "Условие", "Operation", TEP.Choices("operation"), 140); TextColumn(RulesGrid, "Значение", "Value", 170);
            ChoiceColumn(RulesGrid, "Действие", "Action", TEP.Choices("action"), 160); ChoiceColumn(RulesGrid, "Назначение", "Role", TEP.Roles(), 250); ChoiceColumn(RulesGrid, "Часть здания", "Part", TEP.Choices("part"), 180); TextColumn(RulesGrid, "Коэф.", "Coefficient", 80, tip: "Пусто - нормативный коэффициент роли. Число от 0 до 1 заменяет его для площади квартиры с летними помещениями; параметр коэффициента объекта имеет приоритет.");
            ChoiceColumn(CorrectionsGrid, "Показатель", "Metric", MetricChoices, 260); ChoiceColumn(CorrectionsGrid, "Действие", "Action", TEP.Choices("correction"), 210);
            EditableColumn(CorrectionsGrid, "Источник", "Source", Engine.Sources.Select(s => s.Name), 240); TextColumn(CorrectionsGrid, "ID объекта / корпус", "Element", 170);
            TextColumn(CorrectionsGrid, "Новое значение / дельта", "Value", 200); TextColumn(CorrectionsGrid, "Обоснование", "Reason", 360); TextColumn(CorrectionsGrid, "Автор", "Author", 160, true); TextColumn(CorrectionsGrid, "Дата", "Date", 180, true); TextColumn(CorrectionsGrid, "Было", "Previous", 150, true);
            TextColumn(SummaryGrid, "Показатель", "Name", 540, true); TextColumn(SummaryGrid, "Значение", "Value", 135, true, "{0:N" + Config.Decimals + "}"); TextColumn(SummaryGrid, "Ед.", "Unit", 65, true); TextColumn(SummaryGrid, "Статус", "Status", 180, true); TextColumn(SummaryGrid, "Методика", "Method", 170, true); TextColumn(SummaryGrid, "Пояснение", "Comment", 400, true);
            TextColumn(DetailsGrid, "Показатель", "MetricName", 300, true); TextColumn(DetailsGrid, "Корпус", "Building", 145, true); TextColumn(DetailsGrid, "Секция", "Section", 80, true); TextColumn(DetailsGrid, "Этаж", "Level", 155, true); TextColumn(DetailsGrid, "Источник", "Source", 220, true); TextColumn(DetailsGrid, "ElementId", "Element", 100, true);
            TextColumn(DetailsGrid, "Назначение", "PurposeName", 240, true); TextColumn(DetailsGrid, "ID квартиры", "Apartment", 120, true); TextColumn(DetailsGrid, "Исходное", "Raw", 105, true, "{0:N3}"); TextColumn(DetailsGrid, "Коэф.", "Factor", 75, true, "{0:0.###}"); TextColumn(DetailsGrid, "Учтено", "Value", 115, true, "{0:N" + Config.Decimals + "}"); TextColumn(DetailsGrid, "Ед.", "Unit", 65, true);
            CheckColumn(DetailsGrid, "Искл.", "Excluded", 70, true); CheckColumn(DetailsGrid, "Вручную", "Manual", 85, true); TextColumn(DetailsGrid, "Обоснование", "Reason", 480, true);
            TextColumn(IssuesGrid, "Важность", "Severity", 135, true); TextColumn(IssuesGrid, "Код", "Code", 185, true); TextColumn(IssuesGrid, "Сообщение", "Message", 540, true); TextColumn(IssuesGrid, "Показатель", "MetricName", 300, true); TextColumn(IssuesGrid, "Источник", "Source", 220, true); TextColumn(IssuesGrid, "Корпус", "Building", 130, true); TextColumn(IssuesGrid, "ElementId", "Element", 110, true); TextColumn(IssuesGrid, "Что сделать", "Action", 480, true);
        }
        private IEnumerable<T> Children<T>(DependencyObject parent) where T : DependencyObject
        { for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++) { var child = VisualTreeHelper.GetChild(parent, i); if (child is T) yield return (T)child; foreach (var nested in Children<T>(child)) yield return nested; } }
        private void CommitInputs()
        {
            foreach (var grid in new[] { SourcesGrid, BuildingsGrid, LevelsGrid, ParametersGrid, RulesGrid, CorrectionsGrid, MetricSourcesGrid, ContoursGrid })
                if (!grid.CommitEdit(DataGridEditingUnit.Cell, true) || !grid.CommitEdit(DataGridEditingUnit.Row, true)) throw new InvalidOperationException("Исправьте значение в таблице.");
            foreach (var box in Children<TextBox>(this)) box.GetBindingExpression(TextBox.TextProperty)?.UpdateSource();
            if (Children<FrameworkElement>(this).Any(Validation.GetHasError)) throw new InvalidOperationException("Исправьте поля с ошибками ввода.");
            if (Config.Decimals < 0 || Config.Decimals > 6) throw new InvalidOperationException("Точность отображения: от 0 до 6 знаков.");
            foreach (var colour in new[] { Config.IncludedColor, Config.ExcludedColor, Config.ManualColor, Config.PublicColor, Config.SummerColor, Config.TechnicalColor, Config.ErrorColor })
                try { ColorConverter.ConvertFromString(colour); } catch { throw new InvalidOperationException("Неверный цвет: " + colour + ". Используйте формат #RRGGBB."); }
        }
        private void Try(Action action)
        { try { CommitInputs(); action(); } catch (Exception ex) { MessageBox.Show(this, ex.Message, "ТЭП", MessageBoxButton.OK, MessageBoxImage.Warning); Status.Text = ex.Message; } }
        private void Rebind() { DataContext = null; DataContext = this; RestoreContourSelection(); }
        private void RestoreContourSelection()
        { ContourMetric.SelectedValue = Engine.ContourMetricKey; if (ContourMetric.SelectedItem == null) ContourMetric.SelectedIndex = 0; RefreshContourRows(); }
        private void RefreshContourRows()
        { if (ContoursGrid != null) ContoursGrid.ItemsSource = Config.Contours.Where(c => c.Metric == (ContourMetric.SelectedItem as TEP.Metric)?.Key).ToList(); }
        private void ContourMetric_Changed(object sender, SelectionChangedEventArgs e)
        { var metric = ContourMetric.SelectedItem as TEP.Metric; if (metric == null) return; Engine.ContourMetricKey = metric.Key; RefreshContourRows(); }
        private void ContourAction_Click(object sender, RoutedEventArgs e)
        {
            Try(() => {
                if (ContourMetric.SelectedItem == null) throw new InvalidOperationException("Выберите показатель."); var action = (string)((Button)sender).Tag;
                if (action.EndsWith("exclude") && ((TEP.Metric)ContourMetric.SelectedItem).ContourMode == "current") throw new InvalidOperationException("Для ручного выреза сначала выберите способ «По наружным стенам» или «Чертим контур».");
                Engine.PendingContourAction = action; Close();
            });
        }
        private void ClearContourWalls_Click(object sender, RoutedEventArgs e)
        { Try(() => { var m = ContourMetric.SelectedItem as TEP.Metric; if (m == null) return; m.SelectedWalls.Clear(); ContourSettings.DataContext = null; ContourSettings.SetBinding(DataContextProperty, new Binding("SelectedItem") { Source = ContourMetric }); }); }
        private void RemoveContour_Click(object sender, RoutedEventArgs e)
        { Try(() => { foreach (var c in ContoursGrid.SelectedItems.Cast<TEP.ContourSketch>().ToList()) Config.Contours.Remove(c); RefreshContourRows(); }); }
        private void CopyContourSettings(TEP.Metric from, TEP.Metric to)
        {
            if (from == null || to == null || from == to) return;
            to.ContourMode = from.ContourMode; to.WallSelectionMode = from.WallSelectionMode; to.WallParameter = from.WallParameter; to.WallValue = from.WallValue;
            to.WallBoundary = from.WallBoundary; to.WallCutHeight = from.WallCutHeight; to.SelectedWalls = from.SelectedWalls.ToList();
            to.ContourExclusions = from.ContourExclusions; to.VolumeMaskHeight = from.VolumeMaskHeight;
            foreach (var c in Config.Contours.Where(c => c.Metric == to.Key).ToList()) Config.Contours.Remove(c);
            foreach (var c in Config.Contours.Where(c => c.Metric == from.Key).ToList())
                Config.Contours.Add(new TEP.ContourSketch
                {
                    Metric = to.Key,
                    ElementUniqueId = c.ElementUniqueId,
                    ElementLabel = c.ElementLabel,
                    LevelKey = c.LevelKey,
                    LevelName = c.LevelName,
                    Kind = c.Kind,
                    Building = c.Building,
                    Section = c.Section,
                    Profile = c.Profile,
                    BuildingClass = c.BuildingClass,
                    Height = c.Height,
                    BottomOffset = c.BottomOffset
                });
        }
        private void CopyContour_Click(object sender, RoutedEventArgs e)
        { Try(() => { if (ContourCopyFrom.SelectedItem == null) throw new InvalidOperationException("Выберите показатель для копирования."); CopyContourSettings((TEP.Metric)ContourCopyFrom.SelectedItem, (TEP.Metric)ContourMetric.SelectedItem); Rebind(); }); }
        private void AllContours_Click(object sender, RoutedEventArgs e)
        { Try(() => { var from = ContourMetric.SelectedItem as TEP.Metric; foreach (var to in ContourMetrics) CopyContourSettings(from, to); Rebind(); }); }
        private void SelectAll_Click(object sender, RoutedEventArgs e) { changing = true; foreach (var m in Config.Metrics) m.Enabled = true; Rebind(); changing = false; }
        private void SelectNone_Click(object sender, RoutedEventArgs e) { changing = true; foreach (var m in Config.Metrics) m.Enabled = false; Rebind(); changing = false; }
        private void Metric_Checked(object sender, RoutedEventArgs e)
        {
            if (changing || !IsLoaded) return; var check = sender as CheckBox; var metric = check?.DataContext as TEP.Metric; if (metric == null) return;
            var dependencies = new Dictionary<string, string[]> { { "Gns", new[] { "GnsResidential", "GnsNonresidential" } }, { "GnsResidential", new[] { "GnsLivingPart", "GnsNonlivingPart" } }, { "Volume", new[] { "VolumeAbove", "VolumeBelow" } }, { "Gross", new[] { "GrossAbove", "GrossBelow" } }, { "Np", new[] { "NpResidential", "NpNonresidential" } } };
            string[] keys; if (dependencies.TryGetValue(metric.Key, out keys))
            { changing = true; foreach (var key in keys) Config.Metrics.First(x => x.Key == key).Enabled = true; Rebind(); changing = false; }
        }
        private void AddBuilding_Click(object s, RoutedEventArgs e) { Config.Buildings.Add(new TEP.BuildingMap()); }
        private void SectionLevel_Click(object s, RoutedEventArgs e)
        { Try(() => { var level = LevelsGrid.SelectedItem as TEP.LevelSetting; if (level == null) throw new InvalidOperationException("Выберите уровень для уточнения."); var copy = TEP.Engine.Deserialize<TEP.LevelSetting>(TEP.Engine.Serialize(level)); copy.Section = "Укажите секцию"; Config.Levels.Add(copy); LevelsGrid.SelectedItem = copy; LevelsGrid.ScrollIntoView(copy); }); }
        private void DeleteSectionLevel_Click(object s, RoutedEventArgs e)
        { foreach (var level in LevelsGrid.SelectedItems.Cast<TEP.LevelSetting>().Where(x => !string.IsNullOrWhiteSpace(x.Section) || !string.IsNullOrWhiteSpace(x.Building)).ToList()) Config.Levels.Remove(level); }
        private void DeleteBuilding_Click(object s, RoutedEventArgs e) { foreach (var x in BuildingsGrid.SelectedItems.Cast<TEP.BuildingMap>().ToList()) Config.Buildings.Remove(x); }
        private void AddRule_Click(object s, RoutedEventArgs e) { Config.Rules.Add(new TEP.Rule()); }
        private void DeleteRule_Click(object s, RoutedEventArgs e) { foreach (var x in RulesGrid.SelectedItems.Cast<TEP.Rule>().ToList()) Config.Rules.Remove(x); }
        private void AddCorrection_Click(object s, RoutedEventArgs e) { Config.Corrections.Add(new TEP.Correction { Metric = "all", Action = "classify" }); }
        private void DeleteCorrection_Click(object s, RoutedEventArgs e) { foreach (var x in CorrectionsGrid.SelectedItems.Cast<TEP.Correction>().ToList()) Config.Corrections.Remove(x); }
        private void CopyRules_Click(object s, RoutedEventArgs e)
        {
            Try(() => {
                string from = CopyFrom.SelectedValue as string, to = CopyTo.SelectedValue as string; if (from == null || to == null || from == to) throw new InvalidOperationException("Выберите разные исходный и целевой показатели."); int count = 0;
                foreach (var r in Config.Rules.Where(x => x.Metric == from).ToList())
                {
                    var copy = TEP.Engine.Deserialize<TEP.Rule>(TEP.Engine.Serialize(r)); copy.Metric = to;
                    string json = TEP.Engine.Serialize(copy); if (!Config.Rules.Any(x => TEP.Engine.Serialize(x) == json)) { Config.Rules.Add(copy); count++; }
                }
                Status.Text = "Скопировано правил: " + count;
            });
        }
        private void Check_Click(object s, RoutedEventArgs e)
        { Try(() => { SetBusy(true); try { ShowIssues(Engine.CheckParameters(ReportProgress)); Steps.SelectedIndex = 5; ReportTabs.SelectedIndex = 3; ReportStatus.Text = "Проверка наличия параметров и совпадений правил"; Status.Text = "Проверка завершена"; } catch (System.OperationCanceledException) { Status.Text = "Проверка параметров отменена."; } finally { SetBusy(false); } }); }
        private void Save_Click(object s, RoutedEventArgs e) { Try(() => { Engine.SaveSettings(); Status.Text = "Настройки записаны в DataStorage текущего RVT. Сохраните модель, чтобы записать их на диск."; }); }
        private void Import_Click(object s, RoutedEventArgs e)
        { Try(() => { var dialog = new OpenFileDialog { Filter = "Настройки ТЭП (*.json)|*.json", Title = "Импорт настроек" }; if (dialog.ShowDialog(this) != true) return; Engine.ImportSettings(dialog.FileName); Rebind(); ConfigureTables(); Status.Text = "Настройки импортированы; перед расчётом проверьте соответствие источников и параметров."; }); }
        private void ExportSettings_Click(object s, RoutedEventArgs e)
        { Try(() => { var dialog = new SaveFileDialog { Filter = "Настройки ТЭП (*.json)|*.json", FileName = "ТЭП_настройки.json" }; if (dialog.ShowDialog(this) != true) return; Engine.ExportSettings(dialog.FileName); Status.Text = "Настройки экспортированы: " + dialog.FileName; }); }
        private void Calculate_Click(object s, RoutedEventArgs e)
        {
            Try(() =>
            {
                if (!Config.Metrics.Any(x => x.Enabled)) throw new InvalidOperationException("Выберите хотя бы один показатель.");
                SetBusy(true);
                try { var run = Engine.Calculate(ReportProgress); ShowReport(run); Steps.SelectedIndex = 5; Status.Text = "Расчёт завершён. Ошибок: " + run.Issues.Count(x => x.Severity == "Ошибка") + ", предупреждений: " + run.Issues.Count(x => x.Severity == "Предупреждение") + "."; }
                catch (System.OperationCanceledException) { Status.Text = "Расчёт отменён. Изменения этого запуска отменены; предыдущий отчёт сохранён."; }
                finally { SetBusy(false); }
            });
        }
        private void ShowReport(TEP.Run run)
        {
            if (run == null) return; RefreshPrecision(SummaryGrid); RefreshPrecision(DetailsGrid); SummaryGrid.ItemsSource = run.Summary; DetailsGrid.ItemsSource = run.Details; ShowIssues(run.Issues); FloorReport.Document = ColoredReport(run, Config.Decimals);
            string state = run.Summary.Any(x => x.Status == "Неполный результат") ? "Неполный расчёт" : run.Issues.Any(x => x.Severity != "Информация") ? "Расчёт с замечаниями" : "Расчёт завершён";
            ReportStatus.Text = state + " | " + run.Date + " | " + run.Method; ReportStatus.Foreground = run.Issues.Any(x => x.Severity == "Ошибка") ? ErrorTextBrush : (Brush)FindResource("Ink"); ReportTabs.SelectedIndex = 0;
        }
        private static void AppendReportText(Paragraph paragraph, string text, IEnumerable<string> errors)
        {
            // Use complete structured error messages, not code/keyword guesses: a message can contain | and newlines.
            var ranges = new List<Tuple<int, int>>();
            foreach (var error in errors.Where(e => !string.IsNullOrEmpty(e)).Distinct())
                for (int position = 0; (position = text.IndexOf(error, position, StringComparison.Ordinal)) >= 0; position += error.Length)
                    ranges.Add(Tuple.Create(position, position + error.Length));
            int cursor = 0;
            foreach (var range in ranges.OrderBy(r => r.Item1).ThenByDescending(r => r.Item2))
            {
                if (range.Item2 <= cursor) continue;
                if (range.Item1 > cursor) { paragraph.Inlines.Add(new Run(text.Substring(cursor, range.Item1 - cursor))); cursor = range.Item1; }
                paragraph.Inlines.Add(new Run(text.Substring(cursor, range.Item2 - cursor)) { Foreground = ErrorTextBrush }); cursor = range.Item2;
            }
            if (cursor < text.Length) paragraph.Inlines.Add(new Run(text.Substring(cursor)));
        }
        private static FlowDocument ColoredReport(TEP.Run report, int decimals)
        {
            string text = report.TextReport(decimals);
            var paragraph = new Paragraph { Margin = new Thickness(0) };
            var document = new FlowDocument(paragraph) { FontFamily = new FontFamily("Consolas"), FontSize = 13, PagePadding = new Thickness(10), ColumnWidth = double.PositiveInfinity };
            // Preserve the original text, including rounding and the per-floor breakdown. Scope errors to each metric.
            var sections = new List<Tuple<int, TEP.Summary>>(); int search = 0;
            foreach (var summary in report.Summary)
            {
                string header = summary.Name + ": " + Math.Round(summary.Value, decimals, MidpointRounding.AwayFromZero).ToString("N" + decimals) + " " + summary.Unit + " | " + summary.Status;
                int position = text.IndexOf(header, search, StringComparison.Ordinal);
                if (position < 0) continue; sections.Add(Tuple.Create(position, summary)); search = position + header.Length;
            }
            if (sections.Count == 0) { paragraph.Inlines.Add(new Run(text)); return document; }
            paragraph.Inlines.Add(new Run(text.Substring(0, sections[0].Item1)));
            for (int i = 0; i < sections.Count; i++)
            {
                var summary = sections[i].Item2; int start = sections[i].Item1, end = i + 1 < sections.Count ? sections[i + 1].Item1 : text.Length;
                var errors = (report.Issues ?? new List<TEP.Issue>()).Where(e => e.Severity == "Ошибка" && (e.Metric == summary.Key || string.IsNullOrEmpty(e.Metric)))
                    .Select(e => e.Code + ": " + e.Message).ToList();
                if (summary.Status == "Неполный результат") errors.Add("Неполный результат");
                AppendReportText(paragraph, text.Substring(start, end - start), errors);
            }
            return document;
        }
        private void ShowIssues(IEnumerable<TEP.Issue> issues) { displayedIssues = issues.ToList(); RefreshIssues(); }
        private void DiagnosticMode_Changed(object s, SelectionChangedEventArgs e) { RefreshIssues(); }
        private void RefreshIssues()
        {
            if (IssuesGrid == null || displayedIssues == null) return;
            if (Config.IssueDetail != "brief") { IssuesGrid.ItemsSource = displayedIssues; return; }
            IssuesGrid.ItemsSource = displayedIssues.GroupBy(x => new { x.Code, x.Severity, x.Metric, x.Source, x.Building, x.Message }).Select(g => new TEP.Issue
            {
                Code = g.Key.Code,
                Severity = g.Key.Severity,
                Metric = g.Key.Metric,
                Source = g.Key.Source,
                Building = g.Key.Building,
                Message = g.Key.Message + (g.Count() > 1 ? " (повторений: " + g.Count() + ")" : ""),
                Element = g.Count() == 1 ? g.First().Element : "",
                Action = g.First().Action
            }).ToList();
        }
        private void ExportRun_Click(object s, RoutedEventArgs e)
        {
            Try(() => {
                if (Engine.Last == null) throw new InvalidOperationException("Сначала выполните расчёт."); var dialog = new SaveFileDialog { Filter = "Книга Excel (*.xlsx)|*.xlsx|CSV, все разделы (*.csv)|*.csv", FileName = "ТЭП_" + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".xlsx" }; if (dialog.ShowDialog(this) != true) return;
                TEP.Engine.ExportRun(Engine.Last, dialog.FileName, Config.Decimals); Status.Text = "Экспортированы 11 разделов: " + dialog.FileName;
            });
        }
        private void Cleanup_Click(object s, RoutedEventArgs e)
        {
            Try(() => {
                if (MessageBox.Show(this, "Удалить созданные этой версией плагина планы, 3D-виды, цветовые области, расчётные тела и сводные спецификации? Настройки и последний отчёт останутся в RVT.", "Удаление результатов ТЭП", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;
                Status.Text = "Удалено служебных элементов: " + Engine.Cleanup(); if (Engine.Last != null) IssuesGrid.ItemsSource = Engine.Last.Issues.ToList();
            });
        }
        private void Navigate_Click(object s, RoutedEventArgs e)
        {
            TEP.Detail detail = ReportTabs.SelectedIndex == 3 ? null : DetailsGrid.SelectedItem as TEP.Detail;
            if (ReportTabs.SelectedIndex == 3)
            {
                var issue = IssuesGrid.SelectedItem as TEP.Issue;
                if (issue != null) { var source = Engine.Sources.FirstOrDefault(x => x.Name == issue.Source); detail = new TEP.Detail { SourceKey = source?.Key, Source = issue.Source, Element = issue.Element }; }
            }
            if (detail == null) { Status.Text = "Выберите строку объекта или ошибки с ElementId."; return; }
            Engine.RequestedDetail = detail; Close();
        }
        private void Details_DoubleClick(object s, MouseButtonEventArgs e) { Navigate_Click(s, e); }
        private void Issues_DoubleClick(object s, MouseButtonEventArgs e) { Navigate_Click(s, e); }
        private void Back_Click(object s, RoutedEventArgs e) { if (Steps.SelectedIndex > 0) Steps.SelectedIndex--; }
        private void Next_Click(object s, RoutedEventArgs e) { Try(() => { if (Steps.SelectedIndex < Steps.Items.Count - 1) Steps.SelectedIndex++; }); }
        private void Steps_SelectionChanged(object s, SelectionChangedEventArgs e) { if (e.Source == Steps && IsLoaded) UpdateStep(); }
        private void UpdateStep() { BackButton.IsEnabled = !busy && Steps.SelectedIndex > 0; NextButton.IsEnabled = !busy && Steps.SelectedIndex < Steps.Items.Count - 1; }
        private const string Help =
            "1. Подготовьте модель\nРазместите и замкните помещения, пространства или зоны. Проверьте уровни, стадии, параметры назначения, корпуса и номера квартир. Для ручного режима выделите нужные объекты до запуска плагина.\n\n" +
            "2. Выберите расчёт\nНа вкладке «Расчёт» задайте методику ГНС и профиль здания, отметьте необходимые показатели. На вкладке «Контуры» выберите текущий способ, наружные стены или ручные области отдельно для показателя. Для рисования откройте нужный план до запуска плагина; окно временно закроется и вернётся после ввода. Вырезы задаются правилами помещений / элементов и отдельными областями. Для объёма проверьте высоты участков, особенно последнего этажа. Подсказки вариантов объясняют, что меняется в вычислениях.\n\n" +
            "3. Настройте источники\nВключите нужные модели и экземпляры связей. Выберите способ определения корпуса; при необходимости заполните таблицу соответствий и назначьте профили отдельным корпусам. Незагруженную связь загрузите в Revit или исключите.\n\n" +
            "4. Проверьте этажи\nЗадайте ноль здания и землю в координатах основной модели. Отметьте расчётные уровни, вид этажа и наземность. На уклоне добавьте уточнения по корпусу или секции.\n\n" +
            "5. Сопоставьте параметры и правила\nНа вкладке «Правила» выберите источник площадей и схему зон для показателей. Сопоставьте назначения с точными именами параметров. Добавьте условия классификации, включения и исключения. Общие правила действуют первыми, правила показателя уточняют их. Нажмите «Проверить параметры и правила», исправьте обнаруженные проблемы. Примеры ввода находятся в подсказках полей.\n\n" +
            "6. Настройте выдачу\nПо умолчанию создаётся только отчёт в окне. Для визуальной проверки включите «Создавать проверочные виды» над кнопкой расчёта. Тогда плагин дополнительно построит планы с цветными контурами и выбранные 3D-виды / сводную спецификацию. Их параметры задаются на вкладке «Оформление». При выключенном флажке создание и обновление видов полностью пропускаются; вычисление площадей и объёмов выполняется в любом случае.\n\n" +
            "7. Запустите расчёт\nНажмите «Рассчитать ТЭП». Внизу отображаются этап и обрабатываемый объект. Площади обрабатываются плоским алгоритмом; призматические объёмы разделяются по высотам вырезов. Сетка координат 0,001 мм, отклонение хорд дуг до 0,1 мм. Наклонные и другие непризматические тела сохраняют обработку Revit. Версия геометрии указана в заголовке отчёта. Отмена выполняется между операциями Revit и откатывает изменения текущего запуска; отдельная геометрическая операция Revit должна сначала завершиться.\n\n" +
            "8. Проверьте результат\nПросмотрите итоги, площадь каждого этажа, объекты и сообщения об ошибках. Из строки объекта или ошибки перейдите к объекту / виду. Сверьте обводки на служебных планах и 3D. «Неполный результат» требует исправления исходных данных и повторного расчёта.\n\n" +
            "9. Сохраните и перенесите\nЭкспортируйте отчёт в XLSX / CSV. Кнопка «Сохранить в RVT» записывает настройки в служебный DataStorage текущей модели, поле Payload; затем сохраните сам RVT в Revit. JSON переносит настройки между моделями. После импорта заново проверьте источники, параметры и правила. После изменения модели пересчитайте отчёт.";
    }
}
