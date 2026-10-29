using KPLN_Library_Schedules.Common.Schedules;
using KPLN_Library_Schedules.Forms.ViewModels;
using System;
using System.Windows;
using System.Windows.Controls;

namespace KPLN_Library_Schedules.Forms
{
    // Один редактор и одна VM для обоих сценариев. Контекст этажей добавляет только столбцы значений.
    public partial class ScheduleRulesEditor : UserControl
    {
        private ScheduleManagerVM ViewModel => DataContext as ScheduleManagerVM;

        public ScheduleRulesEditor()
        {
            InitializeComponent();
            string constructorHint = RulesHint.Text;
            string constructorSortHint = SortHint.Text;
            object constructorParameterHeader = FilterParameterColumn.Header;
            DataContextChanged += (sender, args) =>
            {
                bool byLevels = ViewModel?.IsByLevels == true;
                var visibility = byLevels ? Visibility.Visible : Visibility.Collapsed;
                OriginalColumn.Visibility = visibility;
                LevelModeColumn.Visibility = visibility;
                FilterParameterColumn.Header = SortParameterColumn.Header = byLevels ? "Параметр" : constructorParameterHeader;
                RulesHint.Text = byLevels
                    ? "Синие строки — существующие правила, зелёные — новые, красные — помеченные к удалению. Параметр можно заменить прямо в строке. Настройте фильтры, порядок сортировки и видимость столбцов; значения фильтров по этажам показаны в нижней таблице."
                    : constructorHint;
                SortHint.Text = byLevels
                    ? "Индекс 1 — главный параметр сортировки. Исходное сохраняет настройки; новое поле по умолчанию скрыто. Видимость одного параметра в обеих таблицах должна совпадать."
                    : constructorSortHint;
            };
        }

        public void Commit()
        {
            foreach (var grid in new[] { BatchGrid, SortGrid })
                if (!grid.CommitEdit(DataGridEditingUnit.Cell, true) || !grid.CommitEdit(DataGridEditingUnit.Row, true))
                    throw new InvalidOperationException("Исправьте значение в редактируемой ячейке.");
        }

        private void RunBusy(Action action, bool commit = true)
        {
            var previous = System.Windows.Input.Mouse.OverrideCursor;
            try
            {
                System.Windows.Input.Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
                if (commit) Commit();
                action();
            }
            catch (Exception ex) { MessageBox.Show(Window.GetWindow(this), ex.Message, "KPLN: Настройки спецификации"); }
            finally { System.Windows.Input.Mouse.OverrideCursor = previous; }
        }
        private void AddRule(object sender, RoutedEventArgs args) { RunBusy(ViewModel.AddRule); }
        private void RemoveRule(object sender, RoutedEventArgs args) { RunBusy(() => ViewModel.RemoveRule(BatchGrid.SelectedItem as BatchScheduleFilter)); }
        private void AddSort(object sender, RoutedEventArgs args) { RunBusy(ViewModel.AddSort); }
        private void RemoveSort(object sender, RoutedEventArgs args) { RunBusy(() => ViewModel.RemoveSort(SortGrid.SelectedItem as BatchScheduleSort)); }
        private void LoadSortParameters(object sender, EventArgs args) { RunBusy(ViewModel.EnsureSortParameters, false); }
        private void LoadFilterParameters(object sender, EventArgs args) { RunBusy(ViewModel.EnsureAvailableParameters, false); }
        private void EditBatchMask(object sender, RoutedEventArgs args)
        {
            var rule = ((FrameworkElement)sender).DataContext as BatchScheduleFilter;
            if (rule == null) return;
            var catalog = ViewModel.IsByLevels ? ScheduleMaskCatalog.CopyByLevels(allowOriginal: rule.IsExisting)
                : ScheduleMaskCatalog.Constructor(allowOriginal: rule.IsExisting);
            var dialog = new ScheduleMaskForm(rule.Value, catalog) { Owner = Window.GetWindow(this) };
            if (dialog.ShowDialog() != true) return;
            rule.Value = dialog.Pattern;
            var levelRule = rule as LevelScheduleFilter;
            if (levelRule != null) levelRule.Mode = FilterModes.Pattern;
        }
    }
}
