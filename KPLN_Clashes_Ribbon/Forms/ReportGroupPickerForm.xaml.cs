using KPLN_Clashes_Ribbon.Core.Reports;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;

namespace KPLN_Clashes_Ribbon.Forms
{
    public partial class ReportGroupPickerForm : Window
    {
        public class ButtonVM
        {
            public int GroupId { get; set; }
            public string Title { get; set; }
            public bool IsEnabled { get; set; }
            public ReportGroup Group { get; set; }
        }

        public ObservableCollection<ButtonVM> Buttons { get; } = new ObservableCollection<ButtonVM>();
        private readonly int _currentGroupId;

        public ReportGroup SelectedGroup { get; private set; }

        public ReportGroupPickerForm(System.Collections.Generic.IEnumerable<ReportGroup> groups, int currentGroupId)
        {
            InitializeComponent();
            DataContext = this;
            _currentGroupId = currentGroupId;

            foreach (var g in groups)
            {
                Buttons.Add(new ButtonVM
                {
                    GroupId = g.Id,
                    Title = $"{g.Name}",            
                    IsEnabled = g.Id != _currentGroupId,
                    Group = g
                });
            }
        }

        private void OnPick(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button b && b.DataContext is ButtonVM vm && vm.IsEnabled)
            {
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
