using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace KPLN_Tools.ExternalCommands.UI
{
    public partial class EvacuationRoutesReportDialog : Window
    {
        private readonly List<EvacuationRoutesStairListItem> _stairs;
        private readonly Action<long> _selectElement;
        private readonly Action _saveReport;
        private readonly Action _openEditor;
        private bool _suppressSelectionAction;

        public EvacuationRoutesStairListItem SelectedStair => DgStairs.SelectedItem as EvacuationRoutesStairListItem;

        public EvacuationRoutesReportDialog(
            IEnumerable<EvacuationRoutesStairListItem> stairs,
            Action<long> selectElement,
            Action saveReport,
            Action openEditor)
        {
            InitializeComponent();

            _stairs = (stairs ?? Enumerable.Empty<EvacuationRoutesStairListItem>()).ToList();
            _selectElement = selectElement;
            _saveReport = saveReport;
            _openEditor = openEditor;

            DgStairs.ItemsSource = _stairs;
            TbHeader.Text = $"Отчёт по лестницам — {_stairs.Count}";
        }

        public void SetBusy(bool busy)
        {
            DgStairs.IsEnabled = !busy;
            BtnSaveReport.IsEnabled = !busy;
            BtnEditor.IsEnabled = !busy;
        }

        public void SetStatus(string text)
        {
            TbStatus.Text = text ?? "";
        }

        public void SelectStair(EvacuationRoutesStairListItem item)
        {
            if (item == null)
                return;

            _suppressSelectionAction = true;
            try
            {
                DgStairs.SelectedItem = item;
                DgStairs.ScrollIntoView(item);
            }
            finally
            {
                _suppressSelectionAction = false;
            }
        }

        private void DgStairs_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_suppressSelectionAction)
                return;

            EvacuationRoutesStairListItem item = SelectedStair;
            if (item == null)
                return;

            _selectElement?.Invoke(item.ElementId);
        }

        private void SaveReport_Click(object sender, RoutedEventArgs e)
        {
            _saveReport?.Invoke();
        }

        private void Editor_Click(object sender, RoutedEventArgs e)
        {
            _openEditor?.Invoke();
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}