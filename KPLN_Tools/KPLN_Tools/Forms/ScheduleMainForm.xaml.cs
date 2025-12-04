using KPLN_Tools.Forms.Models;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace KPLN_Tools.Forms
{
    public partial class ScheduleMainForm : Window
    {
        public ScheduleMainForm(ScheduleFormM model, string vsName)
        {
            InitializeComponent();

            DataContext = new ScheduleFormVM(this, model, vsName);
        }

        private ScheduleFormVM SfVM => DataContext as ScheduleFormVM;

        private void DG_Schedule_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            if (SfVM == null)
                return;

            // Ручной биндинг с обёрткой. Это нужно, если в данных есть символы, которые могут разрушить wpf-разметку (например /)
            if (e.PropertyName is string colName && e.Column is DataGridTextColumn textColumn)
                textColumn.Binding = new System.Windows.Data.Binding($"[{colName}]");
        }

        private void DG_Schedule_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var cell = FindParent<DataGridCell>((DependencyObject)e.OriginalSource);
            if (cell == null)
                return;

            var row = FindParent<DataGridRow>(cell);
            if (row == null)
                return;

            if (!(row.Item is DataRowView rowView))
                return;

            int colIndex = cell.Column.DisplayIndex;

            // Текущее значение
            string currentValue = rowView.Row[colIndex]?.ToString() ?? string.Empty;
            string header = cell.Column.Header?.ToString() ?? string.Empty;

            // Открываем окно для редактирования
            var dlg = new ScheduleSubForm(header, currentValue) { Owner = this };

            bool? result = dlg.ShowDialog();
            if (result == true)
            {
                SfVM?.UpdateCell(rowView, colIndex, dlg.SSFVm);

                if (dlg.SSFVm.AutoIncrement)
                    SfVM?.IncrementColumnBelow(rowView, colIndex, dlg.SSFVm);
            }
        }

        private static T FindParent<T>(DependencyObject child) where T : DependencyObject
        {
            while (child != null)
            {
                if (child is T parent)
                    return parent;

                child = VisualTreeHelper.GetParent(child);
            }
            return null;
        }
    }
}
