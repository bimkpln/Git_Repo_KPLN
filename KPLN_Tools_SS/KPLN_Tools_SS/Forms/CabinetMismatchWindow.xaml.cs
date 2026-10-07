using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools_SS.ExternalCommands;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace KPLN_Tools_SS.Forms
{
    public partial class CabinetMismatchWindow : Window
    {
        private readonly UIApplication _uiapp;

        public CabinetMismatchWindow(
            UIApplication uiapp,
            IList<CabinetMismatchItem> cabinetItems,
            IList<NodeMismatchItem> nodeItems)
        {
            InitializeComponent();
            KPLN_Library_Forms.Services.WindowHandleSearch.MainWindowHandle.SetAsOwner(this);

            _uiapp = uiapp;
            ListBoxCabinets.ItemsSource = cabinetItems;
            ListBoxNodes.ItemsSource = nodeItems;


        }

        private void GoTo_Click(object sender, RoutedEventArgs e)
        {
            Button btn = sender as Button;
            if (btn == null)
                return;

            ElementId id = null;

            CabinetMismatchItem cab = btn.DataContext as CabinetMismatchItem;
            if (cab != null)
            {
                id = cab.ElementId;
            }
            else
            {
                NodeMismatchItem node = btn.DataContext as NodeMismatchItem;
                if (node != null)
                {
                    id = node.ElementId;
                }
            }

            if (id == null)
                return;

            UIDocument uidoc = _uiapp.ActiveUIDocument;
            if (uidoc == null)
                return;

            var ids = new List<ElementId> { id };
            uidoc.Selection.SetElementIds(ids);
            uidoc.ShowElements(id);
        }
    }
}
