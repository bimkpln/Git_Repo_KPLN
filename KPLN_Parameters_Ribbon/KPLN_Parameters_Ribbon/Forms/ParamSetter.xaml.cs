using Autodesk.Revit.DB;
using KPLN_Parameters_Ribbon.Command;
using KPLN_Parameters_Ribbon.Common.CopyElemParamData;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using static KPLN_Loader.Output.Output;

namespace KPLN_Parameters_Ribbon.Forms
{
    public partial class ParamSetter : Window
    {
        private Document Doc { get; set; }

        public ObservableCollection<ListBoxElement> BoxElemsList;

        public ParamSetter(Document doc)
        {
            Doc = doc;
            ObservableCollection<ListBoxElement> heapElemCats = new ObservableCollection<ListBoxElement>();
            HashSet<string> all_ids = new HashSet<string>();
            ObservableCollection<ListBoxElement> all_parameters = new ObservableCollection<ListBoxElement>();
            List<Category> categories = CategoriesList(Doc);
            Array bics = Enum.GetValues(typeof(BuiltInCategory));

            int max = categories.Count;
            string format = "{0} из " + max.ToString() + " категорий проанализировано";

            using (Progress_Single pb = new Progress_Single("KPLN: Копирование параметров", format, max))
            {
                max = 0;
                foreach (Category cat in categories)
                {
                    pb.Increment();
                    try
                    {
                        HashSet<string> ids = new HashSet<string>();
                        ElementId catId = cat.Id;
                        ObservableCollection<ListBoxElement> heapElemParams = new ObservableCollection<ListBoxElement>();

                        FilteredElementCollector typeElemsColl = new FilteredElementCollector(Doc).OfCategoryId(catId).WhereElementIsElementType();
                        FilteredElementCollector instanceElemsColl = new FilteredElementCollector(Doc).OfCategoryId(catId).WhereElementIsNotElementType();
                        if (typeElemsColl.Count() + instanceElemsColl.Count() == 0) { continue; }

                        ListBoxElement lbElement = new ListBoxElement(cat, cat.Name); ;

                        // Обрабатываю параметры типа
                        foreach (ElementType elemType in typeElemsColl)
                        {
                            try
                            {
                                foreach (Parameter p in elemType.Parameters)
                                {
                                    try
                                    {
                                        if (p.StorageType == StorageType.ElementId || p.StorageType == StorageType.None) { continue; }
                                        if (!ids.Contains(p.Definition.Name))
                                        {
                                            ids.Add(p.Definition.Name);
                                            heapElemParams.Add(new ListBoxElement(p, p.Definition.Name, string.Format("{0} : <{1}>", LabelUtils.GetLabelFor(p.Definition.ParameterGroup), p.StorageType.ToString("G"))));
                                        }
                                        if (!all_ids.Contains(p.Definition.Name))
                                        {
                                            all_ids.Add(p.Definition.Name);
                                            all_parameters.Add(new ListBoxElement(p, p.Definition.Name, "Параметры с одним и тем же именем в разных категориях могут отличаться типом данных!"));
                                        }
                                    }
                                    catch (Exception) { }
                                }
                            }
                            catch (Exception) { }
                        }

                        // Обрабатываю параметры экземпляра
                        foreach (Element element in instanceElemsColl)
                        {
                            try
                            {
                                foreach (Parameter p in element.Parameters)
                                {
                                    try
                                    {
                                        if (p.StorageType == StorageType.ElementId || p.StorageType == StorageType.None) { continue; }
                                        if (!ids.Contains(p.Definition.Name))
                                        {
                                            ids.Add(p.Definition.Name);
                                            heapElemParams.Add(new ListBoxElement(p, p.Definition.Name, string.Format("{0} : <{1}>", LabelUtils.GetLabelFor(p.Definition.ParameterGroup), p.StorageType.ToString("G"))));
                                        }
                                        if (!all_ids.Contains(p.Definition.Name))
                                        {
                                            all_ids.Add(p.Definition.Name);
                                            all_parameters.Add(new ListBoxElement(p, p.Definition.Name, "Параметры с одним и тем же именем в разных категориях могут отличаться типом данных!"));
                                        }
                                    }
                                    catch (Exception) { }
                                }
                            }
                            catch (Exception) { }
                        }

                        lbElement.SubElements = new ObservableCollection<ListBoxElement>(heapElemParams.OrderBy(x => x.Name).ThenBy(x => x.ToolTip));
                        heapElemCats.Add(lbElement);
                    }
                    catch (Exception) { }
                }

                BoxElemsList = new ObservableCollection<ListBoxElement>(heapElemCats.OrderBy(x => x.Name));
            }

            InitializeComponent();
            this.RulesControll.ItemsSource = new ObservableCollection<ParameterRuleElement>();
        }

        public void AddRule()
        {
            ParameterRuleElement rule = new ParameterRuleElement(BoxElemsList);
            (this.RulesControll.ItemsSource as ObservableCollection<ParameterRuleElement>).Add(rule);
            UpdateRunEnability();
        }

        private List<Category> CategoriesList(Document doc)
        {
            Categories categories = doc.Settings.Categories;
            List<Category> catList = new List<Category>(categories.Size);


            foreach (Category cat in categories)
            {
                if (!cat.Name.Contains(".dwg") && (cat.CategoryType.Equals(CategoryType.Model) || cat.CategoryType.Equals(CategoryType.Internal)))
                {
                    catList.Add(cat);
                }
            }

            return catList;
        }

        private void OnClickLoadTemplate(object sender, RoutedEventArgs args)
        {
            try
            {
                string path;
                try
                {
                    if (Doc.IsWorkshared && !Doc.IsDetached)
                    {
                        FileInfo info = new FileInfo(ModelPathUtils.ConvertModelPathToUserVisiblePath(Doc.GetWorksharingCentralModelPath()));
                        path = info.Directory.FullName;
                    }
                    else
                    {
                        if (!Doc.IsDetached)
                        {
                            FileInfo info = new FileInfo(Doc.PathName);
                            path = info.Directory.FullName;
                        }
                        else
                        {
                            DirectoryInfo info = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile));
                            path = info.FullName;
                        }

                    }
                }
                catch (Exception)
                {
                    DirectoryInfo info = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile));
                    path = info.FullName;
                }
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Filter = "Text files(*.txt)|*.txt";
                dialog.Title = "Открыть файл настроек";
                dialog.InitialDirectory = path;
                dialog.FileName = "DataRules";
                dialog.ValidateNames = true;
                dialog.DefaultExt = "txt";
                Hide();
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        ParameterRuleElement.LoadData(this, dialog.FileName);
                    }
                    catch (Exception)
                    {
                        EnsureDialog form = new EnsureDialog(this, ":(", "Ошибка при загрузке", "Необходимо проверить правильность выбранного файла с настройкаи.", false);
                        form.ShowDialog();
                    }
                }
                Show();
            }
            catch (Exception e)
            {
                PrintError(e);
                Show();
            }
        }

        private void OnClickSaveAs(object sender, RoutedEventArgs e)
        {
            Save();
        }

        private void OnClickGoToHelp(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start(@"http://moodle.stinproject.local/mod/book/view.php?id=502&chapterid=992");
        }

        private void OnBtnAddRule(object sender, RoutedEventArgs e)
        {
            AddRule();
        }

        private void Save()
        {
            try
            {
                string path;
                try
                {
                    if (Doc.IsWorkshared && !Doc.IsDetached)
                    {
                        FileInfo info = new FileInfo(ModelPathUtils.ConvertModelPathToUserVisiblePath(Doc.GetWorksharingCentralModelPath()));
                        path = info.Directory.FullName;
                    }
                    else
                    {
                        if (!Doc.IsDetached)
                        {
                            FileInfo info = new FileInfo(Doc.PathName);
                            path = info.Directory.FullName;
                        }
                        else
                        {
                            DirectoryInfo info = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile));
                            path = info.FullName;
                        }

                    }
                }
                catch (Exception)
                {
                    DirectoryInfo info = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile));
                    path = info.FullName;
                }
                SaveFileDialog dialog = new SaveFileDialog();
                dialog.Filter = "Text files(*.txt)|*.txt";
                dialog.Title = "Сохранить настройки";
                dialog.InitialDirectory = path;
                dialog.FileName = "DataRules";
                dialog.ValidateNames = true;
                dialog.DefaultExt = "txt";
                Hide();
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    if (!ParameterRuleElement.SaveData(dialog.FileName, this.RulesControll.ItemsSource as ObservableCollection<ParameterRuleElement>))
                    {
                        EnsureDialog form = new EnsureDialog(this, ":(", "Ошибка при сохранении", "При сохранении настроек произошла ошибка...");
                        form.ShowDialog();
                    }
                }
                Show();
            }
            catch (Exception e)
            {
                PrintError(e);
                Show();
            }
        }
        private void OnBtnRun(object sender, RoutedEventArgs e)
        {
            KPLN_Loader.Preferences.CommandQueue.Enqueue(new CommandWriteValues(this.RulesControll.ItemsSource as ObservableCollection<ParameterRuleElement>));
        }

        private void SelectedCategoryChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateRunEnability();
        }

        private void SelectedSourceParamChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateRunEnability();
        }

        private void SelectedTargetParamChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateRunEnability();
        }

        private void UpdateRunEnability()
        {
            if ((this.RulesControll.ItemsSource as ObservableCollection<ParameterRuleElement>).Count == 0)
            {
                BtnRun.IsEnabled = false;
                return;
            }
            foreach (ParameterRuleElement el in this.RulesControll.ItemsSource as ObservableCollection<ParameterRuleElement>)
            {
                if (el.SelectedSourceParameter == null || el.SelectedTargetParameter == null)
                {
                    BtnRun.IsEnabled = false;
                    return;
                }
            }
            BtnRun.IsEnabled = true;
        }

        private void OnBtnRemoveRule(object sender, RoutedEventArgs args)
        {
            IsEnabled = false;
            EnsureDialog form = new EnsureDialog(this, ":O", "Удаление правила", "Необходимо подтверждение");
            form.ShowDialog();
            IsEnabled = true;
            if (EnsureDialog.Commited)
            {
                try
                {
                    (this.RulesControll.ItemsSource as ObservableCollection<ParameterRuleElement>).Remove((sender as System.Windows.Controls.Button).DataContext as ParameterRuleElement);
                }
                catch (Exception e)
                {
                    PrintError(e);
                }
            }
            UpdateRunEnability();
        }
    }
}