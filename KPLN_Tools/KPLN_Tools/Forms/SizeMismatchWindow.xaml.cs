using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.ExternalCommands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace KPLN_Tools.Forms
{
    public partial class SizeMismatchWindow : Window
    {
        private readonly UIApplication _uiapp;

        public SizeMismatchWindow(UIApplication uiapp, IList<SizeMismatchItem> items)
        {
            InitializeComponent();

            _uiapp = uiapp;
            ListBoxItems.ItemsSource = items;

            var helper = new System.Windows.Interop.WindowInteropHelper(this)
            {
                Owner = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle
            };
        }

        private void GoToNode_Click(object sender, RoutedEventArgs e)
        {
            Button btn = sender as Button;
            if (btn == null) return;

            SizeMismatchItem item = btn.DataContext as SizeMismatchItem;
            if (item == null) return;

            if (!item.IsCabinet)
            {
                GoToElement(item.ElementId);
            }
            else
            {
                GoToFirstElementByGroup(isCabinet: false, groupKey: item.GroupKey);
            }
        }

        private void GoToCabinet_Click(object sender, RoutedEventArgs e)
        {
            Button btn = sender as Button;
            if (btn == null) return;

            SizeMismatchItem item = btn.DataContext as SizeMismatchItem;
            if (item == null) return;

            if (item.IsCabinet)
            {
                GoToElement(item.ElementId);
            }
            else
            {
                GoToFirstElementByGroup(isCabinet: true, groupKey: item.GroupKey);
            }
        }

        private void GoToElement(ElementId id)
        {
            UIDocument uidoc = _uiapp.ActiveUIDocument;
            if (uidoc == null) return;

            IList<ElementId> ids = new List<ElementId> { id };
            uidoc.Selection.SetElementIds(ids);
            uidoc.ShowElements(id);
        }

        /// <summary>
        /// Ищет первый ЭлУзел или шкаф с заданным ключом группы и переходит к нему.
        /// </summary>
        private void GoToFirstElementByGroup(bool isCabinet, string groupKey)
        {
            UIDocument uidoc = _uiapp.ActiveUIDocument;
            if (uidoc == null) return;

            Document doc = uidoc.Document;

            string[] nodeFamilies =
            {
                "076_КШ_Шкаф_Универсальный_(ЭлУзл)",
                "076_КШ_Шкаф_Корпус ЩМП У2 IP54_(ЭлУзл)"
            };
            string cabinetFamily = "851_Щит_Универсальный_(ЭлОб)";

            string paramName = isCabinet ? "Имя панели" : "КП_О_Группирование";

            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance));

            foreach (FamilyInstance fi in collector.Cast<FamilyInstance>())
            {
                Family family = fi.Symbol != null ? fi.Symbol.Family : null;
                if (family == null) continue;

                string famName = family.Name;

                if (isCabinet)
                {
                    if (famName != cabinetFamily)
                        continue;
                }
                else
                {
                    if (famName != nodeFamilies[0] && famName != nodeFamilies[1])
                        continue;
                }

                Parameter p = fi.LookupParameter(paramName);
                if (p == null && fi.Symbol != null)
                    p = fi.Symbol.LookupParameter(paramName);

                if (p == null) continue;

                string val = p.AsString();
                if (string.IsNullOrWhiteSpace(val)) continue;

                val = val.Trim();

                if (string.Equals(val, groupKey, StringComparison.OrdinalIgnoreCase))
                {
                    GoToElement(fi.Id);
                    break;
                }
            }
        }
    }
}
