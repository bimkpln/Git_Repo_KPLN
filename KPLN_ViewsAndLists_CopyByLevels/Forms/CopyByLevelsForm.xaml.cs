using KPLN_ViewsAndLists_CopyByLevels.Forms.Entities;
using KPLN_ViewsAndLists_CopyByLevels.Forms.ViewModels;
using KPLN_ViewsAndLists_CopyByLevels.Services;
using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace KPLN_ViewsAndLists_CopyByLevels.Forms
{
    public partial class CopyByLevelsForm : Window
    {
        public CopyByLevelsVM ViewModel { get; }

        public CopyByLevelsForm(CopyByLevelsVM viewModel)
        {
            ViewModel = viewModel;
            InitializeComponent();
            DataContext = viewModel;
        }

        private void CommitGrids()
        {
            foreach (DataGrid grid in new[] { LevelsGrid, ContentsGrid, SchedulesGrid, PreviewGrid })
            {
                grid.CommitEdit(DataGridEditingUnit.Cell, true);
                grid.CommitEdit(DataGridEditingUnit.Row, true);
            }
        }

        private void OptionsChanged(object sender, RoutedEventArgs args)
        {
            ViewModel?.Invalidate();
        }

        private void BuildPreview(object sender, RoutedEventArgs args)
        {
            CommitGrids();
            ViewModel.BuildPreview();
            PreviewGrid.Items.Refresh();
        }

        private void EditFilters(object sender, RoutedEventArgs args)
        {
            CommitGrids();
            var selected = SchedulesGrid.SelectedItem as ScheduleRow;
            if (selected == null) return;
            var dialog = new KPLN_Library_Schedules.Forms.ScheduleFiltersForm(selected, ViewModel.Model.Document, ViewModel.Model.Levels.Where(l => l.Selected)) { Owner = this };
            if (dialog.ShowDialog() == true) ViewModel.Invalidate();
        }

        private void CreateSheets(object sender, RoutedEventArgs args)
        {
            CommitGrids();
            if (!ViewModel.ValidatePreview())
            {
                PreviewGrid.Items.Refresh();
                return;
            }
            try { SettingsService.Save(ViewModel); }
            catch (Exception ex) { MessageBox.Show(this, "Настройки не сохранены: " + ex.Message, "KPLN"); }
            DialogResult = true;
        }
    }
}
