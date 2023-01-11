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
    /// Логика взаимодействия для ElementsToResourcesCompareForm.xaml
    /// </summary>
    public partial class ElementsToResourcesCompareForm : Window
    {
        private static bool Redraw = false;
        public ElementsToResourcesCompareForm()
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
            Redraw = true;
            List<ModelItem> model_items = new List<ModelItem>();
            foreach (DBObject obj in GlobalPreferences.objects)
            {
                model_items.Add(Autodesk.Navisworks.Api.Application.MainDocument.Models.RootItemDescendantsAndSelf.WhereInstanceGuid(obj.Guid).ToList()[0]);
            }
            //
            List<string> categories_list = new List<string>();
            GlobalPreferences.project_categories = GlobalPreferences.GetCategories(model_items);
            foreach (SavedItemCategory cat in GlobalPreferences.project_categories)
            {
                categories_list.Add(cat.DisplayName);
            }
            this.category_picker.ItemsSource = categories_list;
            Redraw = false;
        }

        private void UpdateParameters(string category_name)
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
            this.parameter_picker.ItemsSource = parameter_list;
        }
        private void SelectionChanged(object sender, SelectionChangedEventArgs args)
        {
            if (!Redraw)
            {
                Redraw = true;
                try
                {
                    if (sender == this.category_picker && this.category_picker.SelectedItem.ToString() != "")
                    {
                        UpdateParameters(this.category_picker.SelectedItem.ToString());
                    }
                    try
                    {
                        if (this.parameter_picker.SelectedItem.ToString() != "" && this.category_picker.SelectedItem.ToString() != "")
                        {
                            this.btn_ok.IsEnabled = true;
                        }
                        else
                        {
                            this.btn_ok.IsEnabled = false;
                        }
                    }
                    catch (Exception)
                    {
                        this.btn_ok.IsEnabled = false;
                    }

                }
                catch (Exception)
                {
                }
                Redraw = false;
            }
        }

        private void btn_ok_Click(object sender, RoutedEventArgs args)
        {
            if (this.parameter_picker.SelectedItem.ToString() != "" && this.category_picker.SelectedItem.ToString() != "")
            {
                try
                {
                    this.Hide();
                    Commands.MatchResources(this.category_picker.SelectedItem.ToString(), this.parameter_picker.SelectedItem.ToString());
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
}
