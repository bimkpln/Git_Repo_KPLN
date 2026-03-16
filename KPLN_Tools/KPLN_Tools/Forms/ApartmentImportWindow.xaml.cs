using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;

namespace KPLN_Tools.Forms
{
    public partial class ApartmentImportWindow : Window
    {
        public ObservableCollection<ApartmentImportItemVm> Items { get; private set; }

        public ApartmentImportWindow(List<ApartmentImportItemVm> items)
        {
            InitializeComponent();
            Items = new ObservableCollection<ApartmentImportItemVm>(items);
            DataContext = this;
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}