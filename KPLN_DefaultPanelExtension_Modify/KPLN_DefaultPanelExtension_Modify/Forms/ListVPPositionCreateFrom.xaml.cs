using Autodesk.Revit.DB;
using KPLN_DefaultPanelExtension_Modify.Forms.Models;
using System.Windows;
using System.Windows.Input;

namespace KPLN_DefaultPanelExtension_Modify.Forms
{
    public partial class ListVPPositionCreateFrom : Window
    {
        public ListVPPositionCreateFrom(Window owner, Element ve)
        {
            InitializeComponent();

            CurentListVPPositionCreateVM = new ListVPPositionCreateVM(this, ve);

            Owner = owner;
            DataContext = CurentListVPPositionCreateVM;
        }

        public ListVPPositionCreateVM CurentListVPPositionCreateVM { get; set; }

        private void OnLoaded(object sender, RoutedEventArgs e) => Keyboard.Focus(ConfigName);
    }
}
