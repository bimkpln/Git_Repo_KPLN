using Autodesk.Revit.DB;
using KPLN_Classificator.Forms.ViewModels;
using KPLN_Classificator.Utils;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using static KPLN_Classificator.ApplicationConfig;
using BuiltInCategory = Autodesk.Revit.DB.BuiltInCategory;

namespace KPLN_Classificator.Forms
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public ObservableCollection<RuleItem> ruleItems { get; set; }
        public ClassificatorForm classificatorForm { get; set; }
        public ViewModels.Settings settings { get; set; }
        public InfosStorage storage { get; set; }
        private bool checkedExit { get; set; }

        public MainWindow(ClassificatorForm classificatorForm, InfosStorage storage, List<BuiltInCategory> allCats)
        {
            InitializeComponent();
            this.classificatorForm = classificatorForm;
            this.storage = storage;
            this.ruleItems = new ObservableCollection<RuleItem>();
            this.Collection.ItemsSource = ruleItems;
            this.settings = new ViewModels.Settings();
            this.CollectionParamNames.ItemsSource = settings.paramNameItems;
            RuleItem.builtInCats = allCats;
            this.checkedExit = true;

            if (storage.classificator == null)
            {
                this.storage.classificator = new List<Classificator>();
                this.storage.instanseParams = new List<string>();
                RuleItem firstRuleItem = new RuleItem("", "", "", "", BuiltInCategory.INVALID);
                firstRuleItem.addValueOfParam("");
                ruleItems.Add(firstRuleItem);
                settings.addParamName("Введите имя параметра");
            }
            else
            {
                if (storage.instanceOrType == 1)
                {
                    InstanceRadioButton.IsChecked = true;
                }
                else if (storage.instanceOrType == 2)
                {
                    TypeRadioButton.IsChecked = true;
                }
                foreach (var item in storage.classificator)
                {
                    RuleItem ruleItem = new RuleItem(item.FamilyName, item.TypeName, item.ParameterName, item.ParameterValue, item.BuiltInName);
                    foreach (var pv in item.paramsValues)
                    {
                        ruleItem.addValueOfParam(pv);
                    }
                    ruleItem.setHashSum();
                    ruleItem.colourOfRule = ColourUtils.OLD_RULE;
                    ruleItems.Add(ruleItem);
                }
                foreach (var item in storage.instanseParams)
                {
                    settings.addParamName(item);
                }
            }
            fileName.Text = System.IO.Path.GetFileName(classificatorForm.utils.xmlFilePath);
            rulesCount.Text = ruleItems.Count.ToString();
            try
            {
                lastSaveOfFile.Text = File.GetLastWriteTime(classificatorForm.utils.xmlFilePath).ToString();
            }
            catch (Exception)
            {
                lastSaveOfFile.Text = "Файл не сохранён";
            }
            findDoubledRules();
            refreshNumbersOfRules(ruleItems);
            if (isDocumentAvailable)
            {
                commandEnvironment.toEnqueue(new CommandFindAllElementsInModel(ruleItems));
            }
        }

        public void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            if (checkedExit)
            {
                if (MessageBox.Show("Закрыть окно? Изменения будут потеряны!", "Внимание!", MessageBoxButton.OKCancel, MessageBoxImage.Warning) == MessageBoxResult.OK)
                {
                    classificatorForm.Show();
                }
                else
                {
                    e.Cancel = true;
                }
            }
        }

        private void Delete_Rule_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Удалить правило?", "Внимание!", MessageBoxButton.OKCancel, MessageBoxImage.Information) == MessageBoxResult.OK)
            {
                Button bt = (Button)sender;
                RuleItem ruleItem = bt.DataContext as RuleItem;
                if (ruleItems.Count != 1)
                {
                    ruleItems.Remove(ruleItem);
                    refreshNumbersOfRules(ruleItems);
                }
            }
        }

        private void CopyRule_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Скопировать правило?", "Внимание!", MessageBoxButton.OKCancel, MessageBoxImage.Information) == MessageBoxResult.OK)
            {
                Button bt = (Button)sender;
                RuleItem ruleItem = bt.DataContext as RuleItem;
                RuleItem newRuleItem = new RuleItem(ruleItem.familyNameValue, ruleItem.typeNameValue, ruleItem.parameterName, ruleItem.parameterValue, ruleItem.builtInCategoryName);
                foreach (ParamValueItem item in ruleItem.valuesOfParams)
                {
                    newRuleItem.addValueOfParam(item.paramValue);
                }
                ruleItems.Insert(ruleItems.IndexOf(ruleItem) + 1, newRuleItem);
                refreshNumbersOfRules(ruleItems);
                if (isDocumentAvailable)
                {
                    commandEnvironment.toEnqueue(new CommandFindAllElementsInModel(new ObservableCollection<RuleItem>() { newRuleItem }));
                }
            }
        }

        private void FindElementsByRule_Click(object sender, RoutedEventArgs e)
        {
            if (isDocumentAvailable)
            {
                Button bt = (Button)sender;
                RuleItem ruleItem = bt.DataContext as RuleItem;
                commandEnvironment.toEnqueue(new CommandFindElementsInModel(new List<RuleItem>() { ruleItem }));
            }
            else
            {
                MessageBox.Show("Ни один документ не открыт. Получение информации об элементах невозможна.", "Ошибка!", MessageBoxButton.OKCancel, MessageBoxImage.Warning);
            }
        }

        private void FindAllElementsByRule_Click(object sender, RoutedEventArgs e)
        {
            if (isDocumentAvailable)
            {
                Button bt = (Button)sender;
                RuleItem ruleItem = bt.DataContext as RuleItem;
                commandEnvironment.toEnqueue(new CommandFindElementsInModel(ruleItems.ToList()));
            }
            else
            {
                MessageBox.Show("Ни один документ не открыт. Получение информации об элементах невозможна.", "Ошибка!", MessageBoxButton.OKCancel, MessageBoxImage.Warning);
            }
        }


        private void Add_Rule_Click(object sender, RoutedEventArgs e)
        {
            ruleItems.Add(new RuleItem("", "", "", "", BuiltInCategory.INVALID));
            ruleItems.Last().addValueOfParam("");
            refreshNumbersOfRules(ruleItems);
        }

        private void Get_Rule_Click(object sender, RoutedEventArgs e)
        {
            if (isDocumentAvailable)
            {
                commandEnvironment.toEnqueue(new CommandGetElementInfo(this));
            }
            else
            {
                MessageBox.Show("Ни один документ не открыт. Получение информации об элементах невозможна.", "Ошибка!", MessageBoxButton.OKCancel, MessageBoxImage.Warning);
            }
        }

        public bool setRuleFromElement(BuiltInCategory builtInCategory, string familyName, string typeName, List<string> paramValues)
        {
            if (ruleItems.Count == 1 && ruleItems.First().builtInCategoryName.Equals(BuiltInCategory.INVALID))
            {
                ruleItems.RemoveAt(0);
            }
            RuleItem newRule = new RuleItem(familyName, typeName, "", "", builtInCategory);
            if (!ruleItems.Contains(newRule))
            {
                ruleItems.Add(newRule);
                foreach (string item in paramValues)
                {
                    ruleItems.Last().addValueOfParam(item);
                }
                refreshNumbersOfRules(ruleItems);
                return true;
            }
            return false;
        }

        private void Add_Value_Click(object sender, RoutedEventArgs e)
        {
            Button bt = (Button)sender;
            ParamValueItem valueItem = bt.DataContext as ParamValueItem;
            RuleItem ruleItem = valueItem.parent;
            int index = ruleItem.valuesOfParams.IndexOf(valueItem);
            ruleItem.addValueOfParamByIndex("", index + 1);
        }

        private void Delete_Value_Click(object sender, RoutedEventArgs e)
        {
            Button bt = (Button)sender;
            ParamValueItem valueItem = bt.DataContext as ParamValueItem;
            RuleItem ruleItem = valueItem.parent;
            if (ruleItem.valuesOfParams.Count != 1)
            {
                ruleItem.removeValueOfParam(valueItem);
            }
        }

        private void Add_ParamName_Click(object sender, RoutedEventArgs e)
        {
            Button bt = (Button)sender;
            ParamNameItem paramItem = bt.DataContext as ParamNameItem;
            int index = settings.paramNameItems.IndexOf(paramItem);
            settings.addParamNameByIndex("Введите имя параметра", index + 1);
        }

        private void Delete_ParamName_Click(object sender, RoutedEventArgs e)
        {
            Button bt = (Button)sender;
            ParamNameItem paramNameItem = bt.DataContext as ParamNameItem;
            if (settings.paramNameItems.Count != 1)
            {
                settings.removeParamName(paramNameItem);
            }
        }

        private void Choose_TargetParamName_Click(object sender, RoutedEventArgs e)
        {
            if (isDocumentAvailable)
            {
                Button bt = (Button)sender;
                if (bt.DataContext is ParamNameItem paramNameItem)
                {
                    ParameterSelectorForm parameterSelectorForm = new ParameterSelectorForm(classificatorForm.mparams);
                    parameterSelectorForm.ShowDialog();
                    if (parameterSelectorForm.SelectedMyParameter != null)
                        paramNameItem.paramName = parameterSelectorForm.SelectedMyParameter.Name;
                }
            }
        }

        private void Choose_FilterParamName_Click(object sender, RoutedEventArgs e)
        {
            if (isDocumentAvailable)
            {
                Button bt = (Button)sender;
                if (bt.DataContext is RuleItem ruleItem)
                {
                    ParameterSelectorForm parameterSelectorForm = new ParameterSelectorForm(classificatorForm.mparams);
                    parameterSelectorForm.ShowDialog();
                    if (parameterSelectorForm.SelectedMyParameter != null)
                        ruleItem.parameterName = parameterSelectorForm.SelectedMyParameter.Name;
                }
            }
        }

        private void Insert_ParamName_Click(object sender, RoutedEventArgs e)
        {
            if (isDocumentAvailable)
            {
                Button bt = (Button)sender;
                if (bt.DataContext is ParamValueItem paramValueItem)
                {
                    ParameterSelectorForm parameterSelectorForm = new ParameterSelectorForm(classificatorForm.mparams);
                    parameterSelectorForm.ShowDialog();
                    if (parameterSelectorForm.SelectedMyParameter != null) paramValueItem.paramValue += $"[{parameterSelectorForm.SelectedMyParameter}]";
                }
            }
        }

        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {
            RadioButton pressed = (RadioButton)sender;
            settings.instanceOrType = pressed.Content.ToString().Contains("экземпляр") ? 1 : 2;
        }

        private void refreshNumbersOfRules(ObservableCollection<RuleItem> ruleItems)
        {
            for (int i = 0; i < ruleItems.Count; i++)
            {
                ruleItems[i].ruleNumber = i + 1;
            }
            rulesCount.Text = ruleItems.Count.ToString();
        }

        private void Accept_Classificator_Click(object sender, RoutedEventArgs e)
        {
            if (settings.instanceOrType == 0)
            {
                MessageBox.Show("Выберите тип заполнения классификатора - по типу или по экземпляру.", "Некорректные данные!");
                return;
            }
            if (findDoubledRules())
            {
                if (MessageBox.Show("В файле присутствуют дублирующие друг друга правила! Они подкрашены серым цветом.", "Внимание!", MessageBoxButton.OKCancel, MessageBoxImage.Warning) != MessageBoxResult.OK)
                {
                    return;
                }
            }

            storage.instanceOrType = settings.instanceOrType;
            storage.classificator.Clear();

            foreach (var item in ruleItems)
            {
                Classificator classificator = new Classificator();
                classificator.BuiltInName = item.builtInCategoryName;
                classificator.FamilyName = item.familyNameValue;
                classificator.TypeName = item.typeNameValue;
                classificator.ParameterName = item.parameterName;
                classificator.ParameterValue = item.parameterValue;
                classificator.paramsValues = item.valuesOfParams.Select(x => x.paramValue).ToList();
                storage.classificator.Add(classificator);
            }
            storage.instanseParams.Clear();
            storage.instanseParams.AddRange(settings.paramNameItems.Select(x => x.paramName).ToList());
            checkedExit = false;
            this.Close();
            classificatorForm.Show();
        }

        private void Sort_Classificator_Click(object sender, RoutedEventArgs e)
        {
            ObservableCollection<RuleItem> sortedCollection = new ObservableCollection<RuleItem>();
            foreach (RuleItem elem in ruleItems
                .OrderBy(x => x.builtInCategoryName.ToString())
                .ThenBy(x => x.familyNameValue)
                .ThenBy(x => x.typeNameValue))
            {
                sortedCollection.Add(elem);
            }
            ruleItems = sortedCollection;
            this.Collection.ItemsSource = ruleItems;
            refreshNumbersOfRules(ruleItems);
            findDoubledRules();
            if (isDocumentAvailable)
            {
                commandEnvironment.toEnqueue(new CommandFindAllElementsInModel(ruleItems));
            }
        }

        public bool findDoubledRules()
        {
            IEnumerable<IGrouping<int, RuleItem>> groupingItems = ruleItems
                .GroupBy(ri => ri.GetHashCode())
                .Where(gr => gr.Count() > 1);

            if (groupingItems.Any())
            {
                foreach (IGrouping<int, RuleItem> gr in groupingItems)
                {
                    foreach (RuleItem item in gr)
                    {
                        item.colourOfRule = ColourUtils.DOUBLED_RULE;
                    }
                }
            }

            foreach (RuleItem otherItem in ruleItems)
            {
                bool existsInGroupingItems = groupingItems
                   .SelectMany(gr => gr)
                   .Any(item => item.GetHashCode() == otherItem.GetHashCode());
                if (!existsInGroupingItems && otherItem.colourOfRule.Equals(ColourUtils.DOUBLED_RULE))
                    otherItem.colourOfRule = ColourUtils.NEW_RULE;
            }

            return groupingItems.Any();
        }
    }
}