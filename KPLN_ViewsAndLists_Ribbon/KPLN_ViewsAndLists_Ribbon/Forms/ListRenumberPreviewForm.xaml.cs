using KPLN_ViewsAndLists_Ribbon.Forms.MVVM;
using System.Windows;
using System.Linq;
using System.Windows.Data;
using KPLN_ViewsAndLists_Ribbon.Forms.Entities;
using System.Windows.Controls;

namespace KPLN_ViewsAndLists_Ribbon.Forms
{
    public partial class ListRenumberPreviewForm : Window
    {
        private readonly ListRenumberPreviewVM _viewModel;
        internal System.Action ApplyAction { get; set; }

        private void Back_Click(object sender, RoutedEventArgs e) => Close();

        public ListRenumberPreviewForm(ListRenumberPreviewVM viewModel)
        {
            InitializeComponent();
            _viewModel = viewModel;
            DataContext = viewModel;

            NumberColumn.IsReadOnly = viewModel.IsNumberReadOnly;
            OldParameterColumn.Visibility = viewModel.HasParameter ? Visibility.Visible : Visibility.Collapsed;
            ParameterColumn.Visibility = OldParameterColumn.Visibility;
            ParameterColumn.Header = "Новое значение\n" + viewModel.ParameterName;
        }

        private void ContinueNumbering_Click(object sender, RoutedEventArgs e)
        {
            DataGridCellInfo selectedCell = PreviewGrid.CurrentCell;
            // После перехода фокуса на кнопку CurrentCell может быть пустой,
            // хотя пользовательская ячейка остаётся выделенной.
            if (!selectedCell.IsValid || selectedCell.Column != NumberColumn)
                selectedCell = PreviewGrid.SelectedCells.FirstOrDefault(cell => cell.Column == NumberColumn);

            var selectedRow = selectedCell.IsValid && selectedCell.Column == NumberColumn
                ? selectedCell.Item as ListRenumberPreviewRow : null;

            if (!PreviewGrid.CommitEdit(DataGridEditingUnit.Cell, true)
                || !PreviewGrid.CommitEdit(DataGridEditingUnit.Row, true))
                return;

            // Items содержит видимый порядок, включая сортировку по заголовкам столбцов.
            var orderedRows = PreviewGrid.Items.OfType<ListRenumberPreviewRow>().ToList();
            using (CollectionViewSource.GetDefaultView(PreviewGrid.ItemsSource).DeferRefresh())
                _viewModel.ContinueNumbering(orderedRows, selectedRow);
        }

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            // Последняя редактируемая ячейка тоже должна попасть в применяемый результат.
            if (!PreviewGrid.CommitEdit(DataGridEditingUnit.Cell, true)
                || !PreviewGrid.CommitEdit(DataGridEditingUnit.Row, true))
                return;

            if (_viewModel.CanApply)
                ApplyAction?.Invoke();
        }
    }
}
