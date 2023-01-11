using Autodesk.Navisworks.Api;
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
    /// Логика взаимодействия для CreateSelectionSetsForm.xaml
    /// </summary>
    public partial class CreateSelectionSetsForm : Window
    {
        private static bool Redraw = false;

        public CreateSelectionSetsForm()
        {
            InitializeComponent();
            Loaded += OnLoad;
            Closing += OnClosing;
        }
        
        public void OnClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            GlobalPreferences.state = 0;
        }
        
        private void OnLoad(object sender, RoutedEventArgs e)
        {
            List<string> project_list = new List<string>();
            project_list.Add("<Все элементы>");
            foreach (Model model in Autodesk.Navisworks.Api.Application.ActiveDocument.Models)
            {
                string modelName = model.FileName.Split('\\').Last();
                foreach (ModelItem submodel in model.RootItem.Children)
                {
                    string itemlName = submodel.DisplayName;
                    project_list.Add($"{modelName}: {itemlName}");
                }
            }

            project_list.Sort();
            this.file_picker.ItemsSource = project_list;
        }
        
        private void UpdateCategories(string model_name)
        {
            List<string> categories_list = new List<string>();
            List<ModelItem> model_items;
            if (model_name == "<Все элементы>") 
                model_items = DBTools.GetAllElements(null); 
            else 
                model_items = DBTools.GetAllElements(model_name); 
            
            GlobalPreferences.project_categories = GlobalPreferences.GetCategories(model_items);
            foreach (SavedItemCategory cat in GlobalPreferences.project_categories)
                categories_list.Add(cat.DisplayName);

            categories_list.Sort();
            this.category_picker.ItemsSource = categories_list;
        }

        private void UpdateParameters(string category_name)
        {
            List<string> parameter_list = new List<string>();
            foreach (SavedItemCategory cat in GlobalPreferences.project_categories)
            {
                if (category_name == cat.DisplayName)
                {
                    foreach (string par in cat.HSParameters)
                        parameter_list.Add(par);
                }
            }

            parameter_list.Sort();
            this.parameter_picker.ItemsSource = parameter_list;
        }
        
        private void SelectionChanged(object sender, SelectionChangedEventArgs args)
        {
            if (!Redraw)
            {
                Redraw = true;
                
                if (sender == this.file_picker && this.file_picker.SelectedItem.ToString() != "")
                    UpdateCategories(this.file_picker.SelectedItem.ToString());
                    
                if (sender == this.category_picker && this.category_picker.SelectedItem.ToString() != "")
                    UpdateParameters(this.category_picker.SelectedItem.ToString());
                    
                if (this.parameter_picker.SelectedItem != null && this.category_picker.SelectedItem != null && this.file_picker.SelectedItem != null)
                    this.btn_ok.IsEnabled = true;
                else
                    this.btn_ok.IsEnabled = false;
                
                Redraw = false;
            }
        }

        private void Btn_ok_Click(object sender, RoutedEventArgs args)
        {
            try
            {
                this.Hide();
                    
                if (this.file_picker.SelectedItem.ToString() == "<Все элементы>") 
                    Commands.CreateSelectionSets(null, this.category_picker.SelectedItem.ToString(), this.parameter_picker.SelectedItem.ToString()); 
                else 
                    Commands.CreateSelectionSets(this.file_picker.SelectedItem.ToString(), this.category_picker.SelectedItem.ToString(), this.parameter_picker.SelectedItem.ToString());
                    
                this.Close();  
            }
            catch (Exception e) 
            {
                Output.PrintError(e);
                this.Close(); 
            }
        }
    }
}
