using KPLN_DataBase.Collections;
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

namespace KPLN_Clashes_Ribbon.Forms
{
    /// <summary>
    /// Логика взаимодействия для ProjectPicker.xaml
    /// </summary>
    public partial class ProjectPicker : Window
    {
        public ProjectPicker(Window parent)
        {
            Owner = parent;
            ProjectPickDialog.PickedProject = null;
            InitializeComponent();
            ObservableCollection<DbProject> collection = new ObservableCollection<DbProject>();
            foreach (DbProject p in KPLN_DataBase.DbControll.Projects)
            {
                if (p.Code != "BIM")
                {
                    collection.Add(p);
                }
            }
            Projects.ItemsSource = collection;
            PreviewKeyDown += new KeyEventHandler(HandleEsc);
        }
        private void HandleEsc(object sender, KeyEventArgs e)
        {
            ProjectPickDialog.PickedProject = null;
            if (e.Key == Key.Escape)
            { 
                Close();
            }
        }
        private void OnProjectClick(object sender, RoutedEventArgs e)
        {
            ProjectPickDialog.PickedProject = (sender as Button).DataContext as DbProject;
            Close();
        }
    }
}
