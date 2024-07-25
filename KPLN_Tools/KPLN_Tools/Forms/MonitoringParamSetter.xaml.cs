using Autodesk.Revit.DB;
using KPLN_Library_Forms.Common;
using KPLN_Tools.Common;
using KPLN_Tools.ExecutableCommand;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace KPLN_Tools.Forms
{
    public partial class MonitoringParamSetter : Window
    {
        private readonly Document _doc;
        private readonly Dictionary<ElementId, List<MonitorEntity>> _monitorEntities;
        private readonly ObservableCollection<string> _docParamColl;
        private readonly ObservableCollection<string> _linkParamColl;

        public MonitoringParamSetter(Document doc, Dictionary<ElementId, List<MonitorEntity>> monitorEntities)
        {
            _doc = doc;
            _monitorEntities = monitorEntities;

            ObservableCollection<string> paramColl = new ObservableCollection<string>(_monitorEntities.FirstOrDefault().Value.FirstOrDefault().ModelParameters.Select(p => p.Definition.Name));
            _docParamColl = new ObservableCollection<string>(paramColl.OrderBy(p => p));

            paramColl = new ObservableCollection<string>(_monitorEntities.FirstOrDefault().Value.FirstOrDefault().CurrentMonitorLinkEntity.LinkElemsParams.Select(p => p.Definition.Name));
            _linkParamColl = new ObservableCollection<string>(paramColl.OrderBy(p => p));

            InitializeComponent();
            this.RulesControll.ItemsSource = new ObservableCollection<MonitorParamRule>();
        }

        private void OnBtnAddRule(object sender, RoutedEventArgs e)
        {
            // Создаем новые коллекции для каждого правила
            ObservableCollection<string> docParamCollCopy = new ObservableCollection<string>(_docParamColl);
            ObservableCollection<string> linkParamCollCopy = new ObservableCollection<string>(_linkParamColl);

            MonitorParamRule rule = new MonitorParamRule(docParamCollCopy, linkParamCollCopy);
            (this.RulesControll.ItemsSource as ObservableCollection<MonitorParamRule>).Add(rule);

            UpdateRunEnability();
        }

        private void OnBtnRevalue(object sender, RoutedEventArgs e)
        {
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandExtraMonitoring_SetParams(
                _doc,
                (this.RulesControll.ItemsSource as ObservableCollection<MonitorParamRule>),
                _monitorEntities));
            this.Close();
        }

        private void OnBtnCheck(object sender, RoutedEventArgs e)
        {
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandExtraMonitoring_CheckParams(
               _doc,
               (this.RulesControll.ItemsSource as ObservableCollection<MonitorParamRule>),
               _monitorEntities));
            this.Close();
        }

        private void SelectedSourceParamChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateRunEnability();

            bool isEnabled = false;
            if (sender is ComboBox comboBox)
            {
                if (comboBox.DataContext is MonitorParamRule rule && rule.SelectedSourceParameter != null)
                {
                    if (_linkParamColl.Contains(rule.SelectedSourceParameter))
                        isEnabled = true;
                }
                
                // Получаем контейнер ItemsControl, к которому принадлежит текущий элемент ComboBox
                if (comboBox.TemplatedParent is ContentPresenter contentPresenter)
                {
                    // Находим кнопку BtnRemoveBySource в пределах текущего контекста данных
                    Button btnRemoveBySource = FindChild<Button>(contentPresenter, "BtnRemoveBySource");
                    if (btnRemoveBySource != null)
                    {
                        // Устанавливаем кнопку IsEnabled в true
                        btnRemoveBySource.IsEnabled = isEnabled;
                    }
                }
            }
        }

        private void SelectedTargetParamChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateRunEnability();
        }

        private void UpdateRunEnability()
        {
            if ((this.RulesControll.ItemsSource as ObservableCollection<MonitorParamRule>).Count == 0)
            {
                BtnRevalue.IsEnabled = false;
                BtnCheck.IsEnabled = false;
                return;
            }
            foreach (MonitorParamRule el in this.RulesControll.ItemsSource as ObservableCollection<MonitorParamRule>)
            {
                if (el.SelectedSourceParameter == null || el.SelectedTargetParameter == null)
                {
                    BtnRevalue.IsEnabled = false;
                    BtnCheck.IsEnabled = false;
                    return;
                }
            }
            BtnRevalue.IsEnabled = true;
            BtnCheck.IsEnabled = true;
        }

        private void OnBtnRemoveRule(object sender, RoutedEventArgs args)
        {
            (this.RulesControll.ItemsSource as ObservableCollection<MonitorParamRule>).Remove((sender as System.Windows.Controls.Button).DataContext as MonitorParamRule);

            UpdateRunEnability();
        }

        private void ParamCBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            ComboBox cBox = (ComboBox)sender;
            cBox.IsDropDownOpen = true;

            if (cBox.DataContext is MonitorParamRule rule && rule.SelectedTargetParameter != null)
                cBox.Items.Refresh();
            string searchName = cBox.Text.ToLower();
            
            if (string.IsNullOrWhiteSpace(searchName) == false)
                cBox.Items.Filter = (filterItem) => filterItem.ToString().ToLowerInvariant().Contains(searchName);
            else
            {
                cBox.SelectedItem = null;
                cBox.Items.Filter = (filterItem) => true;
            }

            cBox.Items.Refresh();
        }

        private void BtnRemoveBySource_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button)
            {
                if (button.TemplatedParent is ContentPresenter contentPresenter)
                {
                    if (contentPresenter.DataContext is MonitorParamRule rule && rule.SelectedSourceParameter != null)
                    {
                        if (_linkParamColl.Contains(rule.SelectedSourceParameter))
                            rule.SelectedTargetParameter = rule.SelectedSourceParameter;
                    }
                }
            }
        }

        private T FindChild<T>(DependencyObject parent, string childName) where T : DependencyObject
        {
            // Проверка валидности входных параметров
            if (parent == null) return null;

            T foundChild = null;

            int childrenCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childrenCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (!(child is T))
                {
                    foundChild = FindChild<T>(child, childName);
                    if (foundChild != null) break;
                }
                else if (!string.IsNullOrEmpty(childName))
                {
                    if (child is FrameworkElement frameworkElement && frameworkElement.Name == childName)
                    {
                        foundChild = (T)child;
                        break;
                    }
                }
                else
                {
                    foundChild = (T)child;
                    break;
                }
            }

            return foundChild;
        }
    }
}