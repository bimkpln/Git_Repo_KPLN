using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.ExternalCommands;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using static KPLN_Loader.Output.Output;

namespace KPLN_ModelChecker_User.Forms
{
    /// <summary>
    /// Логика взаимодействия для ElementsOutputExtended.xaml
    /// </summary>
    public partial class ElementsOutputExtended : Window
    {
        public ElementsOutputExtended(ObservableCollection<WPFDisplayItem> collection, ObservableCollection<WPFDisplayItem> categories)
        {
#if Revit2020
            Owner = ModuleData.RevitWindow;
#endif
#if Revit2018
            WindowInteropHelper helper = new WindowInteropHelper(this);
            helper.Owner = ModuleData.MainWindowHandle;
#endif
            InitializeComponent();
            try
            {
                iControll.ItemsSource = collection;
                cbxCategories.ItemsSource = categories;
                cbxCategories.SelectedIndex = 0;
            }
            catch (Exception e)
            {
                PrintError(e);
            }
            
        }
        private void UpdateCollection(int catId)
        {
            foreach (WPFDisplayItem item in iControll.ItemsSource as ObservableCollection<WPFDisplayItem>)
            {
                if (catId == -1)
                {
                    if (item.Visibility != Visibility.Visible)
                    {
                        item.Visibility = Visibility.Visible;
                    }
                }
                else
                {
                    if (item.CategoryId == catId)
                    {
                        if (item.Visibility != Visibility.Visible)
                        {
                            item.Visibility = Visibility.Visible;
                        }
                    }
                    else
                    {
                        if (item.Visibility != Visibility.Collapsed)
                        {
                            item.Visibility = Visibility.Collapsed;
                        }
                    }
                }
            }
        }
        private void OnZoomClick(object sender, RoutedEventArgs e)
        {
            WPFDisplayItem element = (sender as Button).DataContext as WPFDisplayItem;
            element.Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 110, 215, 89));
            if (element.Box != null)
            {
                ModuleData.CommandQueue.Enqueue(new CommandZoomElement(element.Element, element.Box, element.Centroid));
                element.Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 110, 215, 89));
            }
        }

        private void OnSelectedCategoryChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateCollection((cbxCategories.SelectedItem as WPFDisplayItem).CategoryId);
        }
    }
}
