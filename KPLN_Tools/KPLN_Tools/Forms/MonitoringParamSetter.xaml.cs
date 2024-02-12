using Autodesk.Revit.DB;
using KPLN_Tools.Common;
using KPLN_Tools.ExecutableCommand;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace KPLN_Tools.Forms
{
    public partial class MonitoringParamSetter : Window
    {
        private Document _doc { get; set; }

        private Dictionary<ElementId, List<MonitorEntity>> _monitorEntities;

        private ObservableCollection<Parameter> _docParamColl;

        private ObservableCollection<Parameter> _linkParamColl;

        public MonitoringParamSetter(Document doc, Dictionary<ElementId, List<MonitorEntity>> monitorEntities)
        {
            _doc = doc;
            _monitorEntities = monitorEntities;

            ObservableCollection<Parameter> paramColl = new ObservableCollection<Parameter>(_monitorEntities.FirstOrDefault().Value.FirstOrDefault().ModelParameters);
            _docParamColl = new ObservableCollection<Parameter>(paramColl.OrderBy(p => p.Definition.Name));

            paramColl = new ObservableCollection<Parameter>(_monitorEntities.FirstOrDefault().Value.FirstOrDefault().CurrentMonitorLinkEntity.LinkElemsParams);
            _linkParamColl = new ObservableCollection<Parameter>(paramColl.OrderBy(p => p.Definition.Name));

            InitializeComponent();
            this.RulesControll.ItemsSource = new ObservableCollection<MonitorParamRule>();
        }

        private void OnBtnAddRule(object sender, RoutedEventArgs e)
        {
            MonitorParamRule rule = new MonitorParamRule(_docParamColl, _linkParamColl);
            (this.RulesControll.ItemsSource as ObservableCollection<MonitorParamRule>).Add(rule);

            UpdateRunEnability();
        }

        private void OnBtnRevalue(object sender, RoutedEventArgs e)
        {
            KPLN_Loader.Preferences.CommandQueue.Enqueue(new CommandExtraMonitoring_SetParams(
                _doc,
                (this.RulesControll.ItemsSource as ObservableCollection<MonitorParamRule>),
                _monitorEntities));
            this.Close();
        }

        private void OnBtnCheck(object sender, RoutedEventArgs e)
        {
            KPLN_Loader.Preferences.CommandQueue.Enqueue(new CommandExtraMonitoring_CheckParams(
               _doc,
               (this.RulesControll.ItemsSource as ObservableCollection<MonitorParamRule>),
               _monitorEntities));
            this.Close();
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
    }
}