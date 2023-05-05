using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.ExecutableCommand;
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
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using static KPLN_Loader.Output.Output;

namespace KPLN_ModelChecker_User.Forms
{
    /// <summary>
    /// Логика взаимодействия для ElementsOutput.xaml
    /// </summary>
    public partial class ElementsOutput : Window
    {
        public ElementsOutput(ObservableCollection<WPFElement> elements)
        {
            #if Revit2020
            Owner = ModuleData.RevitWindow;
            #endif
            InitializeComponent();
            Items.ItemsSource = elements;
        }

        private void OnZoomElement(object sender, RoutedEventArgs e)
        {
            WPFElement element = (sender as Button).DataContext as WPFElement;
            ModuleData.CommandQueue.Enqueue(new CommandZoomElement(element.Element, element.Box, element.Centroid));
            element.Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 30, 210, 50));
            CollectionViewSource.GetDefaultView(Items.ItemsSource).Refresh();
        }
    }
}
