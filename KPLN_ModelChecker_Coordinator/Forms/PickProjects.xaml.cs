using KPLN_Library_DataBase.Collections;
using KPLN_ModelChecker_Coordinator.DocControll;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Interop;
using static KPLN_Loader.Output.Output;

namespace KPLN_ModelChecker_Coordinator.Forms
{
    /// <summary>
    /// Логика взаимодействия для PickProjects.xaml
    /// </summary>
    public partial class PickProjects : Window
    {
        private static ObservableCollection<DbDocument> DbDocuments { get; set; }
        public PickProjects()
        {
#if Revit2020
            Owner = ModuleData.RevitWindow;
#endif
#if Revit2018
            WindowInteropHelper helper = new WindowInteropHelper(this);
            helper.Owner = ModuleData.MainWindowHandle;
#endif
            InitializeComponent();

            var coll = KPLN_Library_DataBase.DbControll.Projects;
            
            this.lbProjects.ItemsSource = UIWPFElement.GetCollection(KPLN_Library_DataBase.DbControll.Projects);
            
            this.lbSubDepartments.ItemsSource = UIWPFElement.GetCollection(KPLN_Library_DataBase.DbControll.SubDepartments);
            
            this.lbDocuments.ItemsSource = new ObservableCollection<UIWPFElement>();

            IEnumerable<DbDocument> heapDBDocuments = KPLN_Library_DataBase.DbControll.Documents;
            IEnumerable<DbDocument> dbDocuments = new List<DbDocument>();
            foreach (DbDocument dbDocument in heapDBDocuments)
            {
                DbSubDepartment subDep = dbDocument.Department;
                if (subDep != null)
                {
                    int subDepCode = subDep.Id;
                    if (!subDepCode.Equals(22) & !subDepCode.Equals(23) & !subDepCode.Equals(24))
                    {
                        dbDocuments = dbDocuments.Append(dbDocument).ToList();
                    }
                }
            }
            DbDocuments = new ObservableCollection<DbDocument>(dbDocuments);
        }
        
        private bool ProjectIsChecked(DbDocument doc)
        {
            if (doc.Project == null) { return false; }
            foreach (UIWPFElement e in this.lbProjects.ItemsSource as ObservableCollection<UIWPFElement>)
            {
                DbProject project = e.Element as DbProject;
                if (project.Id == doc.Project.Id)
                {
                    if (e.IsChecked)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            return false;
        }
        
        private bool DepartmentIsChecked(DbDocument doc)
        {
            if (doc.Department == null) { return false; }
            foreach (UIWPFElement e in this.lbSubDepartments.ItemsSource as ObservableCollection<UIWPFElement>)
            {
                DbSubDepartment department = e.Element as DbSubDepartment;
                if (department.Id == doc.Department.Id)
                {
                    if (e.IsChecked)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            return false;
        }
        
        private bool ElementInCollection(ObservableCollection<UIWPFElement> collection, DbDocument doc)
        {
            foreach (UIWPFElement e in collection)
            {
                if ((e.Element as DbDocument).Id == doc.Id)
                {
                    return true;
                }
            }
            return false;
        }
        
        private void Update()
        {
            try
            {
                btnStart.IsEnabled = false;
                
                ObservableCollection<UIWPFElement> collection = this.lbDocuments.ItemsSource as ObservableCollection<UIWPFElement>;
                
                if (collection.Count != 0)
                {
                    foreach (UIWPFElement e in collection.ToList())
                    {
                        DbDocument d = e.Element as DbDocument;
                        if (!ProjectIsChecked(d) || !DepartmentIsChecked(d))
                        {
                            collection.Remove(e);
                        }
                        if (e.IsChecked)
                        {
                            btnStart.IsEnabled = true;
                            break;
                        }
                    }
                }
                else
                {
                    foreach (DbDocument d in DbDocuments.OrderBy(db => db.Name))
                    {
                        if (ProjectIsChecked(d) && DepartmentIsChecked(d))
                        {
                            if (!ElementInCollection(collection, d))
                            {
                                collection.Add(new UIWPFElement(d));
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                PrintError(e);
            }

        }
        
        private void OnChecked(object sender, RoutedEventArgs e)
        {
            foreach (UIWPFElement i in this.lbDocuments.SelectedItems)
            {
                i.IsChecked = true;
            }
            Update();
        }

        private void OnUnchecked(object sender, RoutedEventArgs e)
        {
            foreach (UIWPFElement i in this.lbDocuments.SelectedItems)
            {
                i.IsChecked = false;
            }
            Update();
        }

        private void OnCheckedProjects(object sender, RoutedEventArgs e)
        {
            foreach (UIWPFElement i in this.lbProjects.SelectedItems)
            {
                i.IsChecked = true;
            }
            Update();
        }

        private void OnUncheckedProjects(object sender, RoutedEventArgs e)
        {

            foreach (UIWPFElement i in this.lbProjects.SelectedItems)
            {
                i.IsChecked = false;
            }
            Update();
        }

        private void OnUncheckedDepartments(object sender, RoutedEventArgs e)
        {
            foreach (UIWPFElement i in this.lbSubDepartments.SelectedItems)
            {
                i.IsChecked = false;
            }
            Update();
        }
        
        private void OnCheckedDepartments(object sender, RoutedEventArgs e)
        {
            foreach (UIWPFElement i in this.lbSubDepartments.SelectedItems)
            {
                i.IsChecked = true;
            }
            Update();
        }
        
        private void OnBtnStart(object sender, RoutedEventArgs args)
        {
            List<DbDocument> docs = new List<DbDocument>();
            foreach (UIWPFElement e in this.lbDocuments.ItemsSource as ObservableCollection<UIWPFElement>)
            {
                if (e.IsChecked)
                {
                    docs.Add(e.Element as DbDocument);
                }
            }
            KPLN_Loader.Preferences.CommandQueue.Enqueue(new DocChecker(docs));
            Close();
        }
    }
}
