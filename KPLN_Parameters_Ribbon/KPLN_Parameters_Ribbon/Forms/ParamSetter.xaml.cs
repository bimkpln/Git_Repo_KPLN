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
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Parameters_Ribbon.Forms
{
    public partial class ParamSetter : Window
    {
        private readonly Document _doc;
        private readonly HashSet<string> _allIDs = new HashSet<string>();
        private readonly ObservableCollection<ListBoxElement> _heapElemCats = new ObservableCollection<ListBoxElement>();
        private readonly ObservableCollection<ListBoxElement> _allParams = new ObservableCollection<ListBoxElement>();
        private readonly ObservableCollection<ListBoxElement> _boxElemsList;

        public ParamSetter(Document doc)
        {
            _doc = doc;
            List<Category> categories = CategoriesList(_doc);
            Array bics = Enum.GetValues(typeof(BuiltInCategory));

            int max = categories.Count;
            string format = "{0} из " + max.ToString() + " категорий проанализировано";

            using (Progress_Single pb = new Progress_Single("KPLN: Копирование параметров", format, false))
            {
                pb.SetProggresValues(max, 0);
                pb.ShowProgress();

                max = 0;
                foreach (Category cat in categories)
                {
                    pb.Increment();
                    try
                    {
                        HashSet<string> ids = new HashSet<string>();
                        ElementId catId = cat.Id;
                        ObservableCollection<ListBoxElement> heapElemParams = new ObservableCollection<ListBoxElement>();
                        
                        Element[] catElemsColl = new FilteredElementCollector(_doc)
                            .OfCategoryId(catId)
                            .ToArray();
                        if (catElemsColl.Count() == 0) 
                            continue;

                        ListBoxElement lbElement = new ListBoxElement(cat, cat.Name);

                        // Обрабатываю параметры
                        foreach (Element elem in catElemsColl)
                        {
                            AddParamFromElement(elem, ids, heapElemParams);
                            if (elem is ElementType elemType)
                                AddParamFromElement(elemType, ids, heapElemParams);
                        }

                        lbElement.SubElements = new ObservableCollection<ListBoxElement>(heapElemParams.OrderBy(x => x.Name).ThenBy(x => x.ToolTip));
                        _heapElemCats.Add(lbElement);
                    }
                    catch (Exception) { }
                }

                _boxElemsList = new ObservableCollection<ListBoxElement>(_heapElemCats.OrderBy(x => x.Name));
            }

            InitializeComponent();
            this.RulesControll.ItemsSource = new ObservableCollection<ParameterRuleElement>();
        }

        public void AddRule()
        {
            ParameterRuleElement rule = new ParameterRuleElement(_boxElemsList);
            (this.RulesControll.ItemsSource as ObservableCollection<ParameterRuleElement>).Add(rule);
            UpdateRunEnability();
        }

        private void AddParamFromElement(Element elem, HashSet<string> ids, ObservableCollection<ListBoxElement> heapElemParams)
        {
            ParameterSet elemParamSet = elem.Parameters;
            foreach (Parameter param in elemParamSet)
            {
                if (param.StorageType == StorageType.ElementId || param.StorageType == StorageType.None)
                    continue;

                if (!ids.Contains(param.Definition.Name))
                {
                    ids.Add(param.Definition.Name);
                    heapElemParams.Add(new ListBoxElement(
                        param,
                        param.Definition.Name,
                        string.Format("{0} : <{1}>",
#if Revit2020 || Debug2020 || Revit2023 || Debug2023
                        LabelUtils.GetLabelFor(param.Definition.ParameterGroup),
#else
                        LabelUtils.GetLabelForGroup(param.Definition.GetGroupTypeId()),
#endif
                        param.StorageType.ToString("G"))));
                }

                if (!_allIDs.Contains(param.Definition.Name))
                {
                    _allIDs.Add(param.Definition.Name);
                    _allParams.Add(new ListBoxElement(
                        param,
                        param.Definition.Name,
                        "Параметры с одним и тем же именем в разных категориях могут отличаться типом данных!"));
                }
            }
        }

        private List<Category> CategoriesList(Document doc)
        {
            Categories categories = doc.Settings.Categories;
            List<Category> catList = new List<Category>(categories.Size);

            foreach (Category cat in categories)
            {
                if (!cat.Name.Contains(".dwg") && (cat.CategoryType.Equals(CategoryType.Model) || cat.CategoryType.Equals(CategoryType.Internal)))
                    catList.Add(cat);
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
                    if (_doc.IsWorkshared && !_doc.IsDetached)
                    {
                        FileInfo info = new FileInfo(ModelPathUtils.ConvertModelPathToUserVisiblePath(_doc.GetWorksharingCentralModelPath()));
                        path = info.Directory.FullName;
                    }
                    else
                    {
                        if (!_doc.IsDetached)
                        {
                            FileInfo info = new FileInfo(_doc.PathName);
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
                OpenFileDialog dialog = new OpenFileDialog
                {
                    Filter = "Text files(*.txt)|*.txt",
                    Title = "Открыть файл настроек",
                    InitialDirectory = path,
                    FileName = "DataRules",
                    ValidateNames = true,
                    DefaultExt = "txt"
                };
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

        private void OnClickSaveAs(object sender, RoutedEventArgs e) => Save();

        private void OnClickGoToHelp(object sender, RoutedEventArgs e) => System.Diagnostics.Process.Start(@"http://moodle/mod/book/view.php?id=502&chapterid=992");

        private void OnBtnAddRule(object sender, RoutedEventArgs e) => AddRule();

        private void Save()
        {
            try
            {
                string path;
                try
                {
                    if (_doc.IsWorkshared && !_doc.IsDetached)
                    {
                        FileInfo info = new FileInfo(ModelPathUtils.ConvertModelPathToUserVisiblePath(_doc.GetWorksharingCentralModelPath()));
                        path = info.Directory.FullName;
                    }
                    else
                    {
                        if (!_doc.IsDetached)
                        {
                            FileInfo info = new FileInfo(_doc.PathName);
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
                SaveFileDialog dialog = new SaveFileDialog
                {
                    Filter = "Text files(*.txt)|*.txt",
                    Title = "Сохранить настройки",
                    InitialDirectory = path,
                    FileName = "DataRules",
                    ValidateNames = true,
                    DefaultExt = "txt"
                };
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

        private void OnBtnRun_WithoutGroups(object sender, RoutedEventArgs e)
        {
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandWriteValues(this.RulesControll.ItemsSource as ObservableCollection<ParameterRuleElement>, false));
            Close();
        }

        private void OnBtnRun_WithGroups(object sender, RoutedEventArgs e)
        {
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandWriteValues(this.RulesControll.ItemsSource as ObservableCollection<ParameterRuleElement>, true));
            Close();
        }
        
        private void SelectedCategoryChanged(object sender, SelectionChangedEventArgs e) => UpdateRunEnability();

        private void SelectedSourceParamChanged(object sender, SelectionChangedEventArgs e) => UpdateRunEnability();

        private void SelectedTargetParamChanged(object sender, SelectionChangedEventArgs e) => UpdateRunEnability();

        private void UpdateRunEnability()
        {
            if ((this.RulesControll.ItemsSource as ObservableCollection<ParameterRuleElement>).Count == 0)
            {
                BtnRunWithoutGroups.IsEnabled = false;
                BtnRunWithGroups.IsEnabled = false;
                return;
            }
            foreach (ParameterRuleElement el in this.RulesControll.ItemsSource as ObservableCollection<ParameterRuleElement>)
            {
                if (el.SelectedSourceParameter == null || el.SelectedTargetParameter == null)
                {
                    BtnRunWithoutGroups.IsEnabled = false;
                    BtnRunWithGroups.IsEnabled = false;
                    return;
                }
            }
            BtnRunWithoutGroups.IsEnabled = true;
            BtnRunWithGroups.IsEnabled = true;
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