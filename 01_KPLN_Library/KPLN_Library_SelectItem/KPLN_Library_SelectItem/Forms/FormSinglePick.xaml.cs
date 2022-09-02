using KPLN_Library_DataBase.Collections;
using System.Collections;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Linq;
using static KPLN_Loader.Output.Output;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace KPLN_Library_SelectItem.Forms
{
    public partial class FormSinglePick : Window
    {
        private IEnumerable _collection;

        public DbProject SelectedDbProject;

        public FormSinglePick(IEnumerable collection)
        {
            _collection = collection;
            InitializeComponent();

            Projects.ItemsSource = _collection;
            PreviewKeyDown += new KeyEventHandler(HandleEsc);
        }

        private void HandleEsc(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.DialogResult = false;
                Close();
            }
        }

        private void OnElementClick(object sender, RoutedEventArgs e)
        {
            SelectedDbProject = (sender as Button).DataContext as DbProject;
            this.DialogResult = true;
            this.Close();
        }

        /// <summary>
        /// Фильтрация по имени
        /// </summary>
        private void SearchText_Changed(object sender, RoutedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            string _searchName = textBox.Text.ToLower();

            ObservableCollection<DbProject> filteredProjects = new ObservableCollection<DbProject>();

            foreach (DbProject project in _collection)
            {
                if (project.Name.ToLower().StartsWith(_searchName))
                {
                    filteredProjects.Add(project);
                }
            }
            
            Projects.ItemsSource = filteredProjects;
        }

        /// <summary>
        /// Запуск без фильтрации
        /// </summary>
        private void OnRunWithoutFilterClick(object sender, RoutedEventArgs e)
        {
            SelectedDbProject = null;
            this.DialogResult = true;
            this.Close();
        }
    }
}
