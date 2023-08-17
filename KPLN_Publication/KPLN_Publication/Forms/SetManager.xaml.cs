using Autodesk.Revit.DB;
using KPLN_Publication.Common;
using KPLN_Publication.Common.Filters;
using KPLN_Publication.ExternalCommands.PublicationSet;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Publication.Forms
{
    /// <summary>
    /// Логика взаимодействия для SetManager.xaml
    /// </summary>
    public partial class SetManager : Window, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        public void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        private ObservableCollection<ComboBoxSet> _documentSets = new ObservableCollection<ComboBoxSet>();
        public ObservableCollection<ComboBoxSet> DocumentSets
        {
            get
            {
                return _documentSets;
            }
            set
            {
                _documentSets = value;
                NotifyPropertyChanged();
            }
        }
        //
        private Document Doc { get; set; }
        private ObservableCollection<ListBoxParameter> Parameters = new ObservableCollection<ListBoxParameter>();
        private readonly ObservableCollection<string> Types = new ObservableCollection<string>() { "Равно", "Не равно", "Содержит", "Не содержит", "Начиается с", "Заканчивается на" };
        /*
        public void AddSet(Document doc, ViewSheetSet set)
        {
            DocumentSets.Add(new ComboBoxSet(doc, set));
        }
        */
        public SetManager(Document doc)
        {
#if Revit2020
            Owner = ModuleData.RevitWindow;
#endif
#if Revit2018
            WindowInteropHelper helper = new WindowInteropHelper(this);
            helper.Owner = ModuleData.MainWindowHandle;
#endif
            Doc = doc;
            List<ListBoxParameter> parameters = new List<ListBoxParameter>();
            foreach (View v in new FilteredElementCollector(doc).OfClass(typeof(View)).WhereElementIsNotElementType().ToElements())
            {
                if (v.CanBePrinted)
                {
                    foreach (Parameter j in v.Parameters)
                    {
                        if (j.StorageType == StorageType.String && j.HasValue && !ParameterInList(j, parameters))
                        {
                            string toolTip = "";
                            if (v.GetType() == typeof(ViewSheet))
                            { toolTip = "Листы"; }
                            else
                            { toolTip = "Виды"; }
                            parameters.Add(new ListBoxParameter(j, toolTip));
                        }
                    }
                }
            }
            parameters = parameters.OrderBy(x => x.Group).ThenBy(x => x.Name).ThenBy(x => x.ToolTip).ToList();
            Parameters.Add(new ListBoxParameter());
            foreach (ListBoxParameter p in parameters)
            {
                Parameters.Add(p);
            }
            InitializeComponent();
            this.stackPanelFilters.ItemsSource = new ObservableCollection<WPFFilterElement>();
            DocumentSets.Add(new ComboBoxSet(doc, "<Все виды и листы в проекте>", true, true));
            DocumentSets.Add(new ComboBoxSet(doc, "<Все виды>", false, true));
            DocumentSets.Add(new ComboBoxSet(doc, "<Все листы>", true, false));
            foreach (ViewSheetSet set in new FilteredElementCollector(doc).OfClass(typeof(ViewSheetSet)).WhereElementIsNotElementType())
            {
                DocumentSets.Add(new ComboBoxSet(doc, set));
            }
            comboBoxDocumentSets.SelectedIndex = 0;
            DataContext = this;
        }
        private void NumerateFilters()
        {
            //Print("NumerateFilters", KPLN_Loader.Preferences.MessageType.System_Regular);
            int i = 0;
            foreach (WPFFilterElement filter in stackPanelFilters.ItemsSource as ObservableCollection<WPFFilterElement>)
            {
                filter.Number = string.Format("#{0}", (++i).ToString());
            }
        }
        private bool ParameterInList(Parameter p, List<ListBoxParameter> list)
        {
            foreach (var i in list)
            {
                if (p.Id.IntegerValue == i.Parameter.Id.IntegerValue && p.StorageType == i.Parameter.StorageType && p.Definition.Name == i.Parameter.Definition.Name && p.IsShared == i.Parameter.IsShared)
                {
                    return true;
                }
            }
            return false;
        }
        private void OnBtnRemoveFilter(object sender, RoutedEventArgs args)
        {
            try
            {
                WPFFilterElement filterToRemove = (sender as Button).DataContext as WPFFilterElement;
                ObservableCollection<WPFFilterElement> newCollection = new ObservableCollection<WPFFilterElement>();
                ObservableCollection<WPFFilterElement> collection = (stackPanelFilters.ItemsSource as ObservableCollection<WPFFilterElement>);
                foreach (WPFFilterElement el in collection)
                {
                    if (!filterToRemove.Guid.Equals(el.Guid))
                    {
                        newCollection.Add(el);
                    }
                }
                stackPanelFilters.ItemsSource = null;
                stackPanelFilters.ItemsSource = newCollection;
            }
            catch (Exception e)
            {
                PrintError(e);
            }
            NumerateFilters();
            UpdateSheetsVisibility(stackPanelFilters.ItemsSource as ObservableCollection<WPFFilterElement>);
        }
        private void OnBtnAddFilter(object sender, RoutedEventArgs e)
        {
            WPFFilterElement newFilter = new WPFFilterElement();
            (this.stackPanelFilters.ItemsSource as ObservableCollection<WPFFilterElement>).Add(newFilter);
            newFilter.SetTypes(Types);
            newFilter.SetParameters(Parameters);
            NumerateFilters();
        }
        
        private void OnBtnCreateSet(object sender, RoutedEventArgs e)
        {
            List<View> views = new List<View>();
            foreach (ListBoxElement i in listBoxElements.ItemsSource)
            {
                if (i.IsChecked && i.Visibility == System.Windows.Visibility.Visible)
                {
                    views.Add(i.View);
                }
            }
            EnterNameForm form = new EnterNameForm(views, this);
            IsEnabled = false;
            form.ShowDialog();
            IsEnabled = true;
            Refresh();
        }
        
        private void OnBtnRemoveSet(object sender, RoutedEventArgs e)
        {
            ComboBoxSet set = comboBoxDocumentSets.SelectedItem as ComboBoxSet;
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandRemoveSet(set.Set));
            OpenWaitTab();
            comboBoxDocumentSets.SelectedIndex = 0;
            DocumentSets.Remove(set);
            OpenHomeTab();
            Refresh();
        }
        
        public void OpenWaitTab()
        {
            //Print("OpenWaitTab", KPLN_Loader.Preferences.MessageType.System_Regular);
            tb.SelectedIndex = 1;
        }
        
        public void OpenHomeTab()
        {
            //Print("OpenHomeTab", KPLN_Loader.Preferences.MessageType.System_Regular);
            tb.SelectedIndex = 0;
        }
        
        private void UpdateSheetsVisibility(ObservableCollection<WPFFilterElement> filters)
        {
            //Print("UpdateSheetsVisibility", KPLN_Loader.Preferences.MessageType.System_Regular);
            try
            {
                if (DocumentSets == null) { return; }
                if (comboBoxDocumentSets.SelectedItem == null) { return; }
                if (comboBoxDocumentSets.SelectedIndex == -1) { return; }
                if (listBoxElements.ItemsSource as ObservableCollection<ListBoxElement> == null) { return; }
                if ((listBoxElements.ItemsSource as ObservableCollection<ListBoxElement>).Count == 0) { return; }
                foreach (ListBoxElement element in listBoxElements.ItemsSource as ObservableCollection<ListBoxElement>)
                {
                    element.Visibility = System.Windows.Visibility.Visible;
                }
                if (filters == null) { return; }
                if (filters.Count == 0) { return; }
                foreach (ListBoxElement element in listBoxElements.ItemsSource as ObservableCollection<ListBoxElement>)
                {
                    if (element.Visibility == System.Windows.Visibility.Collapsed) { continue; }
                    foreach (WPFFilterElement filter in filters)
                    {
                        try
                        {
                            if (filter.SelectedParameter.Parameter == null || filter.SelectedType == null)
                            {
                                continue;
                            }
                            if (!filter.ElementPassesFilter(element.View))
                            {
                                element.Visibility = System.Windows.Visibility.Collapsed;
                                break;
                            }
                            continue;
                        }
                        catch (Exception)
                        {
                            continue;
                        }

                    }
                    continue;
                }
            }
            catch (Exception e)
            {
                PrintError(e);
            }
        }
        
        private void OnSelectedSetChanged(object sender, SelectionChangedEventArgs e)
        {
            //Print("OnSelectedSetChanged", KPLN_Loader.Preferences.MessageType.System_Regular);
            if (comboBoxDocumentSets.SelectedIndex != -1)
            {
                ComboBoxSet set = comboBoxDocumentSets.SelectedItem as ComboBoxSet;
                ObservableCollection<ListBoxElement> collection = new ObservableCollection<ListBoxElement>();
                foreach (ListBoxElement elem in set.Elements.OrderBy(x => (x.View.GetType() == typeof(ViewSheet)).ToString()).ThenBy(x => x.Name))
                {
                    collection.Add(elem);
                }
                listBoxElements.ItemsSource = collection;
            }
            if (comboBoxDocumentSets.SelectedIndex > 0)
            {
                this.btnRemoveSet.IsEnabled = true;
            }
            else
            {
                this.btnRemoveSet.IsEnabled = false;
            }
            UpdateAddEnability();
            UpdateRemoveEnability();
            UpdateApplyEnability();
            UpdateSheetsVisibility(stackPanelFilters.ItemsSource as ObservableCollection<WPFFilterElement>);
        }
        private void UpdateAddEnability()
        {
            //Print("UpdateAddEnability", KPLN_Loader.Preferences.MessageType.System_Regular);
            foreach (ListBoxElement i in this.listBoxElements.ItemsSource)
            {
                if (i.IsChecked)
                {
                    this.btnAddSet.IsEnabled = true;
                    return;
                }
            }
            this.btnAddSet.IsEnabled = false;
        }
        
        private void UpdateRemoveEnability()
        {
            if (!(comboBoxDocumentSets.SelectedItem is ComboBoxSet set)) { return; }
            
            if (set.IsUserCreated)
            {
                this.btnRemoveSet.IsEnabled = true;
            }
            else
            {
                this.btnRemoveSet.IsEnabled = false;
            }
        }
        
        private void UpdateApplyEnability()
        {
            if (!(comboBoxDocumentSets.SelectedItem is ComboBoxSet set)) { return; }
            
            if (set.IsUserCreated)
            {
                foreach (ListBoxElement i in this.listBoxElements.ItemsSource)
                {
                    if (!i.IsChecked || i.Visibility == System.Windows.Visibility.Collapsed)
                    {
                        this.btnApplyChanges.Visibility = System.Windows.Visibility.Visible;
                        return;
                    }
                }
                this.btnApplyChanges.Visibility = System.Windows.Visibility.Collapsed;
            }
            this.btnApplyChanges.Visibility = System.Windows.Visibility.Collapsed;
        }
        
        private void OnUnchecked(object sender, RoutedEventArgs e)
        {
            //Print("OnUnchecked", KPLN_Loader.Preferences.MessageType.System_Regular);
            UpdateApplyEnability();
            UpdateAddEnability();
            foreach (ListBoxElement i in this.listBoxElements.SelectedItems)
            {
                i.IsChecked = false;
            }
        }
        
        private void OnChecked(object sender, RoutedEventArgs e)
        {
            UpdateApplyEnability();
            UpdateAddEnability();
            foreach (ListBoxElement i in this.listBoxElements.SelectedItems)
            {
                i.IsChecked = true;
            }
        }
        
        public void PickSet(Document doc, ViewSheetSet pickedSet)
        {
            if (comboBoxDocumentSets.SelectedIndex == -1)
            {
                listBoxElements.ItemsSource = null;
            }
            if (pickedSet != null)
            {
                DocumentSets.Add(new ComboBoxSet(doc, pickedSet));
                int i = 0;
                foreach (ComboBoxSet set in DocumentSets)
                {
                    try
                    {
                        if (set.Set.Id.IntegerValue == pickedSet.Id.IntegerValue)
                        {
                            comboBoxDocumentSets.SelectedIndex = i;
                            comboBoxDocumentSets.SelectedItem = set;
                            break;
                        }
                        i++;
                    }
                    catch (Exception)
                    { }
                }
            }
            CollectionViewSource.GetDefaultView(comboBoxDocumentSets.ItemsSource).Refresh();
            CollectionViewSource.GetDefaultView(listBoxElements.ItemsSource).Refresh();
            OpenHomeTab();
        }
        
        private void OnBtnApplyChanges(object sender, RoutedEventArgs e)
        {
            //Print("OnBtnApplyChanges", KPLN_Loader.Preferences.MessageType.System_Regular);
            List<View> views = new List<View>();
            foreach (ListBoxElement i in listBoxElements.ItemsSource)
            {
                if (i.IsChecked && i.Visibility == System.Windows.Visibility.Visible)
                {
                    views.Add(i.View);
                }
            }
            ComboBoxSet set = comboBoxDocumentSets.SelectedItem as ComboBoxSet;
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandApplySet(views, set.Set));
            OpenWaitTab();
            comboBoxDocumentSets.SelectedIndex = 0;
            DocumentSets.Remove(set);
        }
        
        private void OnItemDoubleClick(object sender, MouseButtonEventArgs e)
        {
            //Print("OnItemDoubleClick", KPLN_Loader.Preferences.MessageType.System_Regular);
            try
            {
                View view = (listBoxElements.SelectedItem as ListBoxElement).View;
                if (view != null)
                {
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandSetActiveView(view));
                }
            }
            catch (Exception) { }
        }
        
        private ObservableCollection<string> GetValues(WPFFilterElement filter)
        {
            //Print("GetValues", KPLN_Loader.Preferences.MessageType.System_Regular);
            try
            {
                if (filter.SelectedParameter == null)
                {
                    return new ObservableCollection<string>();
                }
                if (filter.SelectedParameter.Parameter == null)
                {
                    return new ObservableCollection<string>();
                }
                ObservableCollection<string> coll = new ObservableCollection<string>();
                List<string> uniqValues = new List<string>();
                foreach (View view in new FilteredElementCollector(Doc).OfClass(typeof(View)).WhereElementIsNotElementType().ToElements())
                {
                    if (!view.CanBePrinted) { continue; }
                    try
                    {
                        foreach (Parameter p in view.Parameters)
                        {
                            try
                            {
                                if (p.StorageType == StorageType.String && p.IsShared == filter.SelectedParameter.Parameter.IsShared && p.Definition.Name == filter.SelectedParameter.Parameter.Definition.Name)
                                {
                                    string v = p.AsString().ToString();
                                    if (v != string.Empty && !string.IsNullOrWhiteSpace(v) && v != null && !uniqValues.Contains(v))
                                    {
                                        uniqValues.Add(v);
                                    }
                                }
                            }
                            catch (Exception) { }
                        }
                    }
                    catch (Exception ex)
                    { PrintError(ex); }
                }
                uniqValues.Sort();
                foreach (string s in uniqValues)
                {
                    coll.Add(s);
                }
                return coll;
            }
            catch (Exception)
            {
                return new ObservableCollection<string>();
            }
        }
        
        private void OnSelectedParameterChanged(object sender, SelectionChangedEventArgs e)
        {
            //Print("OnSelectedParameterChanged", KPLN_Loader.Preferences.MessageType.System_Regular);
            try
            {
                WPFFilterElement filter = (sender as ComboBox).DataContext as WPFFilterElement;
                filter.Values = GetValues(filter);
                UpdateSheetsVisibility(stackPanelFilters.ItemsSource as ObservableCollection<WPFFilterElement>);
                UpdateApplyEnability();
                UpdateAddEnability();
                return;
            }
            catch (Exception)
            { }
            UpdateApplyEnability();
            UpdateAddEnability();
        }
        
        private void OnSelectedValueChanged(object sender, SelectionChangedEventArgs e)
        {
            //Print("OnSelectedValueChanged", KPLN_Loader.Preferences.MessageType.System_Regular);
            UpdateSheetsVisibility(stackPanelFilters.ItemsSource as ObservableCollection<WPFFilterElement>);
            UpdateApplyEnability();
            UpdateAddEnability();
        }
        
        private void OnSelectedTypeChanged(object sender, SelectionChangedEventArgs e)
        {
            //Print("OnSelectedTypeChanged", KPLN_Loader.Preferences.MessageType.System_Regular);
            UpdateSheetsVisibility(stackPanelFilters.ItemsSource as ObservableCollection<WPFFilterElement>);
            UpdateApplyEnability();
            UpdateAddEnability();
        }
        
        private void Refresh()
        {
            CollectionViewSource.GetDefaultView(comboBoxDocumentSets.ItemsSource).Refresh();
            CollectionViewSource.GetDefaultView(listBoxElements.ItemsSource).Refresh();
        }
    }
}
