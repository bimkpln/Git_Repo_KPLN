using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using TEP = KPLN_Tools.ExternalCommands.Command_AR_CalculateTEP;

namespace KPLN_Tools.Forms
{
    public partial class AR_CalculateTEP : Window
    {
        private readonly TEP.Engine engine;
        private bool ready;
        public AR_CalculateTEP()
        {
            InitializeComponent();
            // WPF sizes are device-independent: keep the initial window within the working area at high DPI.
            var workArea = SystemParameters.WorkArea;
            double availableWidth = Math.Max(640, workArea.Width - 24);
            double availableHeight = Math.Max(480, workArea.Height - 24);
            MinWidth = Math.Min(MinWidth, availableWidth); MinHeight = Math.Min(MinHeight, availableHeight);
            Width = Math.Min(Width, availableWidth); Height = Math.Min(Height, availableHeight);
        }
        public AR_CalculateTEP(TEP.Engine calculationEngine) : this()
        {
            engine = calculationEngine; DataContext = engine.Settings;
            MethodBox.ItemsSource = new[] { "Новая", "Старая" }; MethodBox.SelectedItem = engine.Settings.Method;
            GroupingBox.ItemsSource = new[] { "По связям Revit", "По параметру корпуса", "По рабочему набору", "Ручной выбор" };
            PrimaryBox.ItemsSource = new[] { "Помещения", "Пространства", "Зоны", "Элементы" };
            RoleBox.ItemsSource = TEP.Roles; RoleBox.SelectedIndex = 0;
            OperationBox.ItemsSource = new[] { "Содержит", "Равно", "Не пусто", "Пусто" }; OperationBox.SelectedIndex = 0;
            CategoryBox.ItemsSource = new[] { "" }.Concat(engine.Categories).ToList();
            foreach (var box in new[] { BuildingBox, ParameterBox, ApartmentBox, SectionBox, VoidBox, CoefficientBox }) box.ItemsSource = engine.Parameters;
            LevelBox.ItemsSource = engine.Levels; PhaseBox.ItemsSource = engine.Phases; SchemeBox.ItemsSource = engine.AreaSchemes;
            SourceGrid.ItemsSource = engine.Sources; MethodHelp.Text = TEP.MethodDescription;
            ready = true; RefreshMetrics(0);
        }
        private TEP.Metric Current { get { return MetricList.SelectedItem as TEP.Metric; } }
        private void RefreshMetrics(int index)
        {
            MetricList.ItemsSource = engine.Settings.Metrics; CopySourceBox.ItemsSource = engine.Settings.Metrics;
            MetricList.SelectedIndex = Math.Max(0, index); CopySourceBox.SelectedIndex = 0;
            Editor.DataContext = Current;
        }
        private void MethodChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!ready || MethodBox.SelectedItem == null) return;
            engine.Settings.Method = MethodBox.SelectedItem.ToString(); RefreshMetrics(0);
            StatusText.Text = "Выбрана " + engine.Settings.Method.ToLowerInvariant() + " методика. Настройки другой методики не изменены.";
        }
        private void MetricChanged(object sender, SelectionChangedEventArgs e)
        { if (ready) Editor.DataContext = Current; }
        private void SelectAll(object sender, RoutedEventArgs e) { foreach (var m in engine.Settings.Metrics) m.Enabled = true; MetricList.Items.Refresh(); }
        private void SelectNone(object sender, RoutedEventArgs e) { foreach (var m in engine.Settings.Metrics) m.Enabled = false; MetricList.Items.Refresh(); }
        private void AddRule(object sender, RoutedEventArgs e)
        {
            Guard(() =>
            {
                if (Current == null) return;
                string category = CategoryBox.Text.Trim(), parameter = ParameterBox.Text.Trim();
                if (category.Length == 0 && parameter.Length == 0) throw new InvalidOperationException("Выберите категорию или укажите имя параметра.");
                if (category.Length > 0 && !engine.Categories.Contains(category)) throw new InvalidOperationException("Категория не найдена в модели. Выберите её из списка.");
                string operation = OperationBox.SelectedItem.ToString(), role = RoleBox.SelectedItem.ToString();
                if (parameter.Length > 0 && (operation == "Равно" || operation == "Содержит") && string.IsNullOrWhiteSpace(ValueBox.Text))
                    throw new InvalidOperationException("Введите значение; для пустого параметра используйте условие «Пусто».");
                double factor = 1;
                if (role == "Летнее помещение" && (!TEP.TryNumber(FactorBox.Text, out factor) || factor < 0 || factor > 1))
                    throw new InvalidOperationException("Коэффициент должен быть числом от 0 до 1.");
                var rule = new TEP.Rule { Role = role, Category = category, Parameter = parameter, Operation = operation, Value = ValueBox.Text.Trim(), Factor = factor };
                if (Current.Rules.Any(r => r.Label == rule.Label)) throw new InvalidOperationException("Такое правило уже добавлено.");
                Current.Rules.Add(rule); StatusText.Text = "Правило добавлено: " + rule.Label;
            });
        }
        private void RemoveRule(object sender, RoutedEventArgs e)
        { var rule = ((FrameworkElement)sender).DataContext as TEP.Rule; if (Current != null && rule != null) Current.Rules.Remove(rule); }
        private void FindValues(object sender, RoutedEventArgs e)
        {
            Guard(() => { var values = engine.Values(CategoryBox.Text.Trim(), ParameterBox.Text.Trim()); ValueBox.ItemsSource = values; StatusText.Text = "Найдено различных значений: " + values.Count; });
        }
        private void CheckRules(object sender, RoutedEventArgs e)
        { Guard(() => { DetailsBox.Text = engine.CheckParameters(Current); ReportTabs.SelectedIndex = 1; Tabs.SelectedIndex = 2; StatusText.Text = "Проверка параметров выполнена. Модель не изменена."; }); }
        private void CopySettings(object sender, RoutedEventArgs e)
        {
            if (Current == null || CopySourceBox.SelectedItem == null) return;
            int index = Current.Index; bool enabled = Current.Enabled;
            var copy = ((TEP.Metric)CopySourceBox.SelectedItem).Copy(index); copy.Enabled = enabled;
            engine.Settings.Metrics[index] = copy; RefreshMetrics(index);
            StatusText.Text = "Настройки скопированы. Проверьте правила назначения для текущего показателя.";
        }
        private void ApplyToAll(object sender, RoutedEventArgs e)
        {
            if (Current == null) return;
            var source = Current.Copy(Current.Index); int selected = Current.Index;
            for (int i = 0; i < engine.Settings.Metrics.Count; i++)
            { bool enabled = engine.Settings.Metrics[i].Enabled; var copy = source.Copy(i); copy.Enabled = enabled; engine.Settings.Metrics[i] = copy; }
            RefreshMetrics(selected); StatusText.Text = "Настройки применены ко всем показателям текущей методики. Уточните различия классификации.";
        }
        private void ReadGround(object sender, RoutedEventArgs e)
        { Guard(() => { StatusText.Text = engine.GroundFromSelection(); GroundBox.GetBindingExpression(TextBox.TextProperty).UpdateTarget(); }); }
        private void SaveSettings(object sender, RoutedEventArgs e)
        { Guard(() => { CommitEditors(); engine.SaveSettings(); StatusText.Text = "Настройки записаны в текущую модель. Сохраните RVT (Ctrl+S), чтобы они сохранились после закрытия проекта."; }); }
        private void Calculate(object sender, RoutedEventArgs e) { RunCalculation(false); }
        private void CalculateGraphics(object sender, RoutedEventArgs e) { RunCalculation(true); }
        private void CommitEditors()
        {
            Keyboard.ClearFocus(); SourceGrid.CommitEdit(DataGridEditingUnit.Cell, true); SourceGrid.CommitEdit(DataGridEditingUnit.Row, true);
        }
        private void RunCalculation(bool graphics)
        {
            Guard(() =>
            {
                CommitEditors();
                double ground; if (!TEP.TryNumber(engine.Settings.Ground, out ground)) throw new InvalidOperationException("Отметка земли должна быть числом в метрах.");
                StatusText.Text = "Расчёт…"; Dispatcher.Invoke(DispatcherPriority.Render, new Action(() => { }));
                var run = engine.Calculate();
                if (graphics)
                {
                    try { engine.CreateGraphics(); }
                    catch (Exception ex) { run.GraphicsStatus = "ОШИБКА ГРАФИКИ: " + ex.Message; }
                }
                ReportBox.Text = run.Report(); DetailsBox.Text = run.Report(true);
                ReportTabs.SelectedIndex = 0; Tabs.SelectedIndex = 2;
                StatusText.Text = "Расчёт завершён. Неполных показателей: " + run.Results.Count(r => r.Incomplete) +
                    "; требуют проверки: " + run.Results.Count(r => r.NeedsReview) + ". Настройки сохраняются отдельной кнопкой.";
            });
        }
        private void Guard(Action action)
        {
            try { Mouse.OverrideCursor = Cursors.Wait; action(); }
            catch (Exception ex) { StatusText.Text = "Ошибка: " + ex.Message; MessageBox.Show(this, ex.Message, "Расчёт ТЭП", MessageBoxButton.OK, MessageBoxImage.Warning); }
            finally { Mouse.OverrideCursor = null; }
        }
        private void CloseWindow(object sender, RoutedEventArgs e) { Close(); }
    }
}
