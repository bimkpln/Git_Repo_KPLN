using KPLN_Library_Schedules.Common.Schedules;
using KPLN_Library_Schedules.Forms.ViewModels;
using System;
using System.Windows;
using System.Windows.Controls;

namespace KPLN_Library_Schedules.Forms
{
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


        private void Commit()
        {
            RulesEditor.Commit();
            foreach (DataGrid grid in new[] { SchedulesGrid, ValuesGrid })
            {
                if (!grid.CommitEdit(DataGridEditingUnit.Cell, true) || !grid.CommitEdit(DataGridEditingUnit.Row, true))
                    throw new InvalidOperationException("Исправьте значение в редактируемой ячейке.");
            }
        }

        private void SelectAll(object sender, RoutedEventArgs args) { ViewModel.SelectSchedules(true); }
        private void SelectNone(object sender, RoutedEventArgs args) { ViewModel.SelectSchedules(false); }
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
            var window = ScheduleReportForm.Create(ViewModel.ExecutionReport);
            window.Owner = this;
            window.ShowDialog();
        }
    }
}
