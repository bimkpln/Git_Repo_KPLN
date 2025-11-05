using KPLN_Clashes_Ribbon.Core.Reports;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace KPLN_Clashes_Ribbon.Forms
{
    public partial class ReportGroupPickerForm : Window
    {
        public class ButtonVM
        {
            public int GroupId { get; set; }
            public string Title { get; set; }
            public string NameRaw { get; set; }   
            public bool IsEnabled { get; set; }
            public ReportGroup Group { get; set; }
        }

        public ObservableCollection<ButtonVM> Buttons { get; } = new ObservableCollection<ButtonVM>();

        private readonly int _currentGroupId;
        private ICollectionView _buttonsView;
        private string _search = string.Empty;

        public ReportGroup SelectedGroup { get; private set; }
        public int SelectedNumber { get; private set; }  

        public ReportGroupPickerForm(IEnumerable<ReportGroup> groups, int currentGroupId)
        {
            InitializeComponent();
            DataContext = this;
            _currentGroupId = currentGroupId;

            foreach (var g in groups)
            {
                Buttons.Add(new ButtonVM
                {
                    GroupId = g.Id,
                    Title = (g.Name ?? string.Empty) + " (ID: " + g.Id + ")",
                    NameRaw = g.Name ?? string.Empty,
                    IsEnabled = g.Id != _currentGroupId, 
                    Group = g
                });
            }

            _buttonsView = CollectionViewSource.GetDefaultView(Buttons);
            _buttonsView.Filter = ButtonsFilter;

            ThresholdTB.Text = "5";
            SelectedNumber = 0;

            UpdateEmptyState();
        }


        /// Фильтр: скрыть текущую группу; поиск по имени группы (без регистра).
        private bool ButtonsFilter(object obj)
        {
            var vm = obj as ButtonVM;
            if (vm == null) return false;

            if (vm.GroupId == _currentGroupId)
                return false;

            if (string.IsNullOrWhiteSpace(_search))
                return true;

            return vm.NameRaw != null &&
                   vm.NameRaw.IndexOf(_search, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void OnSearchTextChanged(object sender, TextChangedEventArgs e)
        {
            _search = (SearchTB.Text ?? string.Empty).Trim();
            if (_buttonsView != null) _buttonsView.Refresh();
            UpdateEmptyState();
        }

        private void OnClearSearch(object sender, RoutedEventArgs e)
        {
            SearchTB.Text = string.Empty;
            SearchTB.Focus();
        }

        private void UpdateEmptyState()
        {
            bool isEmpty = _buttonsView == null || _buttonsView.IsEmpty;
            ButtonsList.Visibility = isEmpty ? Visibility.Collapsed : Visibility.Visible;
            EmptyLabel.Visibility = isEmpty ? Visibility.Visible : Visibility.Collapsed;
        }

        private void OnThresholdPreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            for (int i = 0; i < e.Text.Length; i++)
            {
                if (!char.IsDigit(e.Text[i]))
                {
                    e.Handled = true;
                    return;
                }
            }
        }

        private void OnThresholdPasting(object sender, DataObjectPastingEventArgs e)
        {
            if (e.DataObject.GetDataPresent(typeof(string)))
            {
                var text = e.DataObject.GetData(typeof(string)) as string;
                if (!IsDigitsOnly(text))
                    e.CancelCommand();
            }
            else
            {
                e.CancelCommand();
            }
        }

        private static bool IsDigitsOnly(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            for (int i = 0; i < s.Length; i++)
                if (!char.IsDigit(s[i])) return false;
            return true;
        }

        private void OnThresholdLostFocus(object sender, RoutedEventArgs e)
        {
            int val;
            if (!int.TryParse(ThresholdTB.Text, out val))
                val = 0;

            if (val < 0) val = 0;
            if (val > 300) val = 300;

            ThresholdTB.Text = val.ToString();
            SelectedNumber = val;
        }

        private void OnPick(object sender, RoutedEventArgs e)
        {
            var b = sender as Button;
            if (b == null) return;

            var vm = b.DataContext as ButtonVM;
            if (vm != null && vm.IsEnabled)
            {
                OnThresholdLostFocus(ThresholdTB, null);

                SelectedGroup = vm.Group;
                DialogResult = true;
                Close();
            }
        }

        private void OnCancel(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
