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
    /// Логика взаимодействия для CreateQuantItemsForm.xaml
    /// </summary>
    public partial class CreateQuantItemsForm : Window
    {
        private static bool Redraw = false;
        public CreateQuantItemsForm()
        {
            InitializeComponent();
            Loaded += OnLoad;
            Closing += OnClosing;
        }
        public void OnClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            GlobalPreferences.state = 0;
        }
        public void OnLoad(object sender, RoutedEventArgs e)
        {
            //List<string> value_collection = new List<string>();
            //foreach (SavedItem item in Autodesk.Navisworks.Api.Application.ActiveDocument.SelectionSets.RootItem.Children)
            //{
            //    value_collection.Add(item.DisplayName);
            //}
            List<string> value_collection = APITools.GetAllSaveditemsNames();
            value_collection.Sort();
            this.file_picker.ItemsSource = value_collection;
        }
        private void UpdateCategories(string model_name)
        {
            List<string> categories_list = new List<string>();
            List<ModelItem> model_items = new List<ModelItem>();
            //List<SavedItem> saved_items = APITools.GetSavedItems(model_name);
            List<SavedItem> saved_items = new List<SavedItem>();

            List<SavedItem> saved_item_groups = APITools.GetSavedItemGroups();
            foreach (SavedItem group in saved_item_groups)
            {
                if (group.DisplayName == model_name)
                {
                    foreach (SavedItem child_item in APITools.GetSubItems(group))
                    {
                        saved_items.Add(child_item);
                    }
                }
            }
            foreach (SavedItem item in saved_items)
            {
                ModelItemCollection model_collection = ((SelectionSet)item).GetSelectedItems();
                foreach (ModelItem model_item in model_collection)
                {
                    model_items.Add(model_item);
                }
            }
            GlobalPreferences.project_categories = GlobalPreferences.GetCategories(model_items);
            foreach (SavedItemCategory cat in GlobalPreferences.project_categories)
            {
                if (cat.HSParameters.Count != 0)
                { 
                    categories_list.Add(cat.DisplayName); 
                }
                
            }
            
            categories_list.Sort();
            this.category_picker1.ItemsSource = categories_list;
            this.category_picker2.ItemsSource = categories_list;
        }
        private void UpdateParameters(string category_name, ComboBox target)
        {
            List<string> parameter_list = new List<string>();
            foreach (SavedItemCategory cat in GlobalPreferences.project_categories)
            {
                if (category_name == cat.DisplayName)
                {
                    foreach (string par in cat.HSParameters)
                    {
                        parameter_list.Add(par);
                    }                    
                }
            }

            parameter_list.Sort();
            target.ItemsSource = parameter_list;
        }
        private void SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!Redraw)
            {
                Redraw = true;
                if (sender == this.file_picker)
                {
                    UpdateCategories(this.file_picker.SelectedItem.ToString());
                }
                if (sender == this.category_picker1)
                {
                    UpdateParameters(this.category_picker1.SelectedItem.ToString(), this.parameter_picker1);
                }
                if (sender == this.category_picker2)
                {
                    UpdateParameters(this.category_picker2.SelectedItem.ToString(), this.parameter_picker2);
                }

                if (this.parameter_picker2.SelectedItem != null && this.category_picker2.SelectedItem != null && this.parameter_picker1.SelectedItem != null)
                    this.btn_ok.IsEnabled = true;
                else
                    this.btn_ok.IsEnabled = false;
                
                Redraw = false;
            }
        }

        private void btn_ok_Click(object sender, RoutedEventArgs args)
        {
            try
            {
                this.Hide();
                Commands.AddQuantItems(this.file_picker.Text.ToString(), this.category_picker1.SelectedItem.ToString(), this.parameter_picker1.SelectedItem.ToString(), this.category_picker2.SelectedItem.ToString(), this.parameter_picker2.SelectedItem.ToString());
                //Commands.AddQuantItems("ОС.ЗП.2.5", "Тип в приложении Revit", "Описание по классификатору", "Тип в приложении Revit", "Описание");
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
