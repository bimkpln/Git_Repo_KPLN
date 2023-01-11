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

namespace KPLN_Quantificator.Forms
{
    /// <summary>
    /// Логика взаимодействия для WPFPickFolder.xaml
    /// </summary>
    public partial class ElemtsToQuantItemsForm : Window
    {
        public ElemtsToQuantItemsForm()
        {
            InitializeComponent();
            Loaded += OnLoad;
            Closing += OnClosing;
        }
        public void OnClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            GlobalPreferences.state = 0;
        }
        private void OnLoad(object sender, RoutedEventArgs args)
        {
            List<string> value_collection = APITools.GetAllSaveditemsNames();
            this.cbx_folders.ItemsSource = value_collection;
        }
        private void btn_ok_Click(object sender, RoutedEventArgs args)
        {
            try
            {
                this.Hide();
                //Commands.UpdateQuantification(this.cbx_folders.Text, this.remove_previous.IsChecked.Value);
                Commands.UpdateQuantification("ОС.ЗП.2.5", this.remove_previous.IsChecked.Value);
                this.Close();
            }
            catch (Exception e)
            {
                Output.PrintError(e);
                this.Close();

            }
            
        }

        private void Selected_Index_Changed(object sender, SelectionChangedEventArgs e)
        {
            //try
            //{
            //    this.btn_ok.IsEnabled = true;
            //}
            //catch (Exception)
            //{
            //    this.btn_ok.IsEnabled = false;
            //}
            this.btn_ok.IsEnabled = true;
        }
    }
}
