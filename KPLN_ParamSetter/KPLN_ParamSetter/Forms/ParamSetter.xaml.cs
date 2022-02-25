using Autodesk.Revit.DB;
using KPLN_ParamSetter.Command;
using KPLN_ParamSetter.Common;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using static KPLN_Loader.Output.Output;

namespace KPLN_ParamSetter.Forms
{
    /// <summary>
    /// Логика взаимодействия для ParamSetter.xaml
    /// </summary>
    public partial class ParamSetter : Window
    {
        private Document Doc { get; set; }
        public ObservableCollection<ListBoxElement> _Categories = new ObservableCollection<ListBoxElement>();
        public ParamSetter(Document doc)
        {
#if Revit2020
            Owner = ModuleData.RevitWindow;
#endif
#if Revit2018
            WindowInteropHelper helper = new WindowInteropHelper(this);
            helper.Owner = ModuleData.MainWindowHandle;
#endif
            _Categories.Add(new ListBoxElement(null, "<Нет>"));
            ObservableCollection<ListBoxElement> _сategories = new ObservableCollection<ListBoxElement>();
            Doc = doc;
            HashSet<string> all_ids = new HashSet<string>();
            ObservableCollection<ListBoxElement> all_parameters = new ObservableCollection<ListBoxElement>();
            InitializeComponent();
            this.RulesControll.ItemsSource = new ObservableCollection<ParameterRuleElement>();
            foreach (BuiltInCategory built_in_category in Enum.GetValues(typeof(BuiltInCategory)))
            {
                
                try
                {
                    Category cat = Category.GetCategory(Doc, built_in_category);
                    HashSet<string> ids = new HashSet<string>();
                    if (new FilteredElementCollector(Doc).OfCategory(built_in_category).WhereElementIsElementType().Count() + new FilteredElementCollector(Doc).OfCategoryId(cat.Id).WhereElementIsNotElementType().Count() == 0)
                    {
                        continue;
                    }
                    ObservableCollection<ListBoxElement> parameters = new ObservableCollection<ListBoxElement>();
                    ListBoxElement category = new ListBoxElement(cat, cat.Name);
                    if (new FilteredElementCollector(Doc).OfCategory(built_in_category).WhereElementIsElementType().Count() != 0)
                    {
                        foreach (ElementType type in new FilteredElementCollector(Doc).OfCategory(built_in_category).WhereElementIsElementType().ToElements())
                        {
                            try
                            {
                                if (type != null)
                                {
                                    foreach (Parameter p in type.Parameters)
                                    {
                                        try
                                        {
                                            if (p.StorageType == StorageType.ElementId || p.StorageType == StorageType.None) { continue; }
                                            if (!ids.Contains(p.Definition.Name))
                                            {
                                                ids.Add(p.Definition.Name);
                                                parameters.Add(new ListBoxElement(p, p.Definition.Name, string.Format("{0} : <{1}>", LabelUtils.GetLabelFor(p.Definition.ParameterGroup), p.StorageType.ToString("G"))));
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
                            }
                            catch (Exception) { }
                        }
                    }
                    if (new FilteredElementCollector(Doc).OfCategory(built_in_category).WhereElementIsNotElementType().Count() != 0)
                    {
                        foreach (Element element in new FilteredElementCollector(Doc).OfCategory(built_in_category).WhereElementIsNotElementType().ToElements())
                        {
                            try
                            {
                                if (element != null)
                                {
                                    foreach (Parameter p in element.Parameters)
                                    {
                                        try
                                        {
                                            if (p.StorageType == StorageType.ElementId || p.StorageType == StorageType.None) { continue; }
                                            if (!ids.Contains(p.Definition.Name))
                                            {
                                                ids.Add(p.Definition.Name);
                                                parameters.Add(new ListBoxElement(p, p.Definition.Name, string.Format("{0} : <{1}>", LabelUtils.GetLabelFor(p.Definition.ParameterGroup), p.StorageType.ToString("G"))));
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
                            }
                            catch (Exception) { }
                        }
                    }
                    ObservableCollection<ListBoxElement> sortedParameters = new ObservableCollection<ListBoxElement>();
                    foreach (ListBoxElement lbParameter in parameters.OrderBy(x => x.Name).ThenBy(x => x.ToolTip).ToList())
                    {
                        sortedParameters.Add(lbParameter);
                    }
                    category.SubElements = sortedParameters;
                    _сategories.Add(category);
                }
                catch (Exception) { }
            }
            ListBoxElement all_cats = new ListBoxElement(null, "<Все категории>");
            all_cats.SubElements = new ObservableCollection<ListBoxElement>();
            foreach (ListBoxElement par in all_parameters)
            { all_cats.SubElements.Add(par); }
            foreach (ListBoxElement lbCategory in _сategories.OrderBy(x => x.Name).ToList())
            {
                _Categories.Add(lbCategory);
            }
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

        }

        private void OnBtnAddRule(object sender, RoutedEventArgs e)
        {
            AddRule();
        }
        public void AddRule()
        {
            ObservableCollection<ListBoxElement> _сategories = new ObservableCollection<ListBoxElement>();
            foreach (ListBoxElement i in _Categories) { _сategories.Add(i); }
            ParameterRuleElement rule = new ParameterRuleElement(_сategories);
            (this.RulesControll.ItemsSource as ObservableCollection<ParameterRuleElement>).Add(rule);
            UpdateRunEnability();
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
            try
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
            catch (Exception)
            { }
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
