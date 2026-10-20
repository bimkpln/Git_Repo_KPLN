using KPLN_ViewsAndLists_Ribbon.Common.Schedules;
using KPLN_ViewsAndLists_Ribbon.Forms.ViewModels;
using System;
using System.Windows;
using System.Windows.Controls;

namespace KPLN_ViewsAndLists_Ribbon.Forms
{
    // Поиск внутри выпадающего списка не изменяет ItemsSource и выбранный параметр.
    public sealed class ScheduleParameterComboBox : ComboBox
    {
        private TextBox _search;
        public ScheduleParameterComboBox()
        {
            IsTextSearchEnabled = false;
            VirtualizingStackPanel.SetIsVirtualizing(this, false);
            ItemContainerGenerator.StatusChanged += (s, e) => FilterItems();
        }

        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();
            var popup = GetTemplateChild("PART_Popup") as System.Windows.Controls.Primitives.Popup;
            if (popup == null || popup.Child == null) return;
            var content = popup.Child;
            popup.Child = null;
            var panel = new DockPanel { LastChildFill = true, Background = System.Windows.Media.Brushes.White };
            _search = new TextBox { Margin = new Thickness(4), Padding = new Thickness(4),
                ToolTip = "Поиск по любой части имени параметра. Очистите строку, чтобы показать все параметры." };
            DockPanel.SetDock(_search, Dock.Top);
            panel.Children.Add(_search);
            panel.Children.Add(content);
            popup.Child = panel;
            _search.TextChanged += (s, e) => FilterItems();
            _search.PreviewKeyDown += (s, e) =>
            {
                if (e.Key == System.Windows.Input.Key.Escape)
                {
                    IsDropDownOpen = false;
                    Focus();
                    e.Handled = true;
                }
                else if (e.Key == System.Windows.Input.Key.Down)
                {
                    for (int i = 0; i < Items.Count; i++)
                    {
                        var item = ItemContainerGenerator.ContainerFromIndex(i) as ComboBoxItem;
                        if (item != null && item.Visibility == Visibility.Visible) { item.Focus(); break; }
                    }
                    e.Handled = true;
                }
            };
        }

        protected override void OnDropDownOpened(EventArgs e)
        {
            base.OnDropDownOpened(e);
            if (_search == null) return;
            _search.Text = "";
            Dispatcher.BeginInvoke(new Action(() => { FilterItems(); _search.Focus(); }),
                System.Windows.Threading.DispatcherPriority.Input);
        }

        private void FilterItems()
        {
            string query = (_search?.Text ?? "").Trim();
            for (int i = 0; i < Items.Count; i++)
            {
                var container = ItemContainerGenerator.ContainerFromIndex(i) as ComboBoxItem;
                if (container == null) continue;
                var parameter = Items[i] as ScheduleParameterChoice;
                bool visible = query.Length == 0 || (parameter?.Name ?? "").IndexOf(query, StringComparison.CurrentCultureIgnoreCase) >= 0;
                container.Visibility = visible ? Visibility.Visible : Visibility.Collapsed;
                container.IsEnabled = visible;
            }
        }
    }

    public partial class ScheduleManagerForm : Window
    {
        public ScheduleManagerVM ViewModel { get; }
        public ScheduleManagerForm(ScheduleManagerVM viewModel)
        {
            ViewModel = viewModel;
            InitializeComponent();
            DataContext = viewModel;
        }

        private void EditNameMask(object sender, RoutedEventArgs args)
        {
            var dialog = new ScheduleMaskForm(ViewModel.NamePattern,
                ScheduleMaskCatalog.Constructor(true, ViewModel.IsParameterCopy)) { Owner = this };
            if (dialog.ShowDialog() == true) ViewModel.NamePattern = dialog.Pattern;
        }

        private void EditBatchMask(object sender, RoutedEventArgs args)
        {
            var rule = ((FrameworkElement)sender).DataContext as BatchScheduleFilter;
            if (rule == null) return;

            var dialog = new ScheduleMaskForm(rule.Value,
                ScheduleMaskCatalog.Constructor(allowOriginal: rule.IsExisting)) { Owner = this };
            if (dialog.ShowDialog() == true) rule.Value = dialog.Pattern;
        }

        private void Commit()
        {
            foreach (DataGrid grid in new[] { SchedulesGrid, ValuesGrid, BatchGrid })
            {
                if (!grid.CommitEdit(DataGridEditingUnit.Cell, true) || !grid.CommitEdit(DataGridEditingUnit.Row, true))
                    throw new InvalidOperationException("Исправьте значение в редактируемой ячейке.");
            }
        }

        private void SelectAll(object sender, RoutedEventArgs args) { ViewModel.SelectSchedules(true); }
        private void SelectNone(object sender, RoutedEventArgs args) { ViewModel.SelectSchedules(false); }
        private void AddRule(object sender, RoutedEventArgs args)
        {
            var previousCursor = System.Windows.Input.Mouse.OverrideCursor;
            try
            {
                System.Windows.Input.Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
                ViewModel.AddRule();
            }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, "KPLN: Конструктор спецификаций"); }
            finally { System.Windows.Input.Mouse.OverrideCursor = previousCursor; }
        }
        private bool _loadingParameters;
        private void LoadCopyParameters(object sender, EventArgs args)
        {
            _loadingParameters = true;
            try { RunBusy(ViewModel.EnsureAvailableParameters); }
            finally { _loadingParameters = false; }
        }
        private void CopyParameterChanged(object sender, SelectionChangedEventArgs args)
        {
            if (!_loadingParameters && ViewModel.IsParameterCopy && ViewModel.CopyParameter != null)
                RunBusy(ViewModel.ReadCopyValues);
        }
        private void ReadCopyValues(object sender, RoutedEventArgs args) { RunBusy(ViewModel.ReadCopyValues); }
        private void SelectAllValues(object sender, RoutedEventArgs args) { ViewModel.SelectValues(true); }
        private void SelectNoValues(object sender, RoutedEventArgs args) { ViewModel.SelectValues(false); }
        private void RunBusy(Action action)
        {
            var previous = System.Windows.Input.Mouse.OverrideCursor;
            try { System.Windows.Input.Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait; action(); }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, Title); }
            finally { System.Windows.Input.Mouse.OverrideCursor = previous; }
        }

        private void RemoveRule(object sender, RoutedEventArgs args)
        {
            var rule = BatchGrid.SelectedItem as BatchScheduleFilter;
            ViewModel.RemoveRule(rule);
        }

        private void Preview(object sender, RoutedEventArgs args)
        {
            RunBusy(() =>
            {
                Commit();
                ViewModel.BuildPreview();
                if (ViewModel.Operations.Count > 0) Tabs.SelectedIndex = 2;
                else MessageBox.Show(this, ViewModel.Status, Title);
            });
        }

        private void Execute(object sender, RoutedEventArgs args)
        {
            bool executed = false;
            RunBusy(() =>
            {
                Commit();
                if (!ViewModel.Ready) return;
                ViewModel.Execute();
                executed = true;
                SchedulesGrid.Items.Refresh();
                Tabs.SelectedIndex = 2;
            });
            if (executed) ShowExecutionReport(sender, args);
        }

        private void ShowExecutionReport(object sender, RoutedEventArgs args)
        {
            if (ViewModel.ExecutionReport == null) return;
            new ScheduleReportForm(ViewModel.ExecutionReport) { Owner = this }.ShowDialog();
        }
    }
}
