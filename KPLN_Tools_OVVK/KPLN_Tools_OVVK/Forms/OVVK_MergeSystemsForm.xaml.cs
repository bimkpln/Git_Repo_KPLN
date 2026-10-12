using Autodesk.Revit.DB;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UIFactory;
using KPLN_Tools_OVVK.Common.OVVK_System;
using KPLN_Tools_OVVK.ExecutableCommand;
using KPLN_Tools_OVVK.Forms.Entities;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_Tools_OVVK.Forms
{
    public partial class OVVK_MergeSystemsForm : Window
    {
        private readonly OVVK_SystemManagerM _model;

        public OVVK_MergeSystemsForm(OVVK_SystemManagerM model)
        {
            _model = model;

            InitializeComponent();
            SysIControll.ItemsSource = SysDataToMerge;

            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        public ObservableCollection<OVVK_MergeSystem> SysDataToMerge { get; set; } = new ObservableCollection<OVVK_MergeSystem>();
        
        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
        }

        private void AddNewGroup_Click(object sender, RoutedEventArgs e)
        {
            SysDataToMerge.Add(new OVVK_MergeSystem(_model.SystemSumParameters.ToArray()));
        }

        private void Run_Click(object sender, RoutedEventArgs e)
        {
            if (!SysDataToMerge.Any())
                return;

            if (SysDataToMerge.Any(group => string.IsNullOrWhiteSpace(group.FirstParamName)
                || string.IsNullOrWhiteSpace(group.LastParamName)
                || group.FirstParamName == group.LastParamName))
            {
                MessageBox.Show(this, "В каждой группе выбери две разные системы.",
                    "KPLN: Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string[] names = SysDataToMerge.SelectMany(group => new[] { group.FirstParamName, group.LastParamName }).ToArray();
            if (names.Distinct().Count() != names.Length)
            {
                MessageBox.Show(this, "Одна система может участвовать только в одной группе объединения за запуск.",
                    "KPLN: Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            DialogResult = true;
        }

        private void Del_Click(object sender, RoutedEventArgs e)
        {
            OVVK_MergeSystem ent = (sender as Button).DataContext as OVVK_MergeSystem;
            SysDataToMerge.Remove(ent);
        }
    }
}
