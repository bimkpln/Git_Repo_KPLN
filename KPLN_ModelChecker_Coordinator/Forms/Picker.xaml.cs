using KPLN_Library_DataBase.Collections;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;

namespace KPLN_ModelChecker_Coordinator.Forms
{
    public partial class Picker : Window
    {
        public static DbProject PickedProject = null;
        public static DbDocument PickedDocument = null;
        public Picker(List<DbProject> projects)
        {
#if Revit2020
            Owner = ModuleData.RevitWindow;
#endif
#if Revit2018
            WindowInteropHelper helper = new WindowInteropHelper(this);
            helper.Owner = ModuleData.MainWindowHandle;
#endif
            PickedProject = null;
            PickedDocument = null;
            InitializeComponent();
            this.Projects.ItemsSource = projects;
            tbHeader.Text = "Проекты:";
            Title = "KPLN: Выбрать проект";
        }
        public Picker(List<DbDocument> documents)
        {
#if Revit2020
            Owner = ModuleData.RevitWindow;
#endif
#if Revit2018
            WindowInteropHelper helper = new WindowInteropHelper(this);
            helper.Owner = ModuleData.MainWindowHandle;
#endif
            PickedProject = null;
            PickedDocument = null;
            InitializeComponent();
            this.Projects.ItemsSource = documents;
            tbHeader.Text = "Документы:";
            Title = "KPLN: Выбрать документ";
        }
        private void OnBtnClick(object sender, RoutedEventArgs e)
        {
            if ((sender as Button).DataContext.GetType() == typeof(DbDocument))
            {
                PickedDocument = (sender as Button).DataContext as DbDocument;
            }
            if ((sender as Button).DataContext.GetType() == typeof(DbProject))
            {
                PickedProject = (sender as Button).DataContext as DbProject;
            }
            Close();
        }
    }
}
