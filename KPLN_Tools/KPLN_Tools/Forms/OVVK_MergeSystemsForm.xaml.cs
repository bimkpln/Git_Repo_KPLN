using Autodesk.Revit.DB;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UIFactory;
using KPLN_Tools.Common.OVVK_System;
using KPLN_Tools.ExecutableCommand;
using KPLN_Tools.Forms.Models;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_Tools.Forms
{
    public partial class OVVK_MergeSystemsForm : Window
    {
        private readonly OVVK_SystemManager_VM _ovvk_ViewModel;

        public OVVK_MergeSystemsForm(OVVK_SystemManager_VM ovvk_ViewModel)
        {
            _ovvk_ViewModel = ovvk_ViewModel;

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
            SysDataToMerge.Add(new OVVK_MergeSystem(_ovvk_ViewModel.SystemSumParameters.Where(ssp => !ssp.Contains("ВНИМАНИЕ!!!")).ToArray()));
        }

        private void Run_Click(object sender, RoutedEventArgs e)
        {
            if (SysDataToMerge.Any()) 
                this.DialogResult = true;

            Close();
        }

        private void Del_Click(object sender, RoutedEventArgs e)
        {
            OVVK_MergeSystem ent = (sender as Button).DataContext as OVVK_MergeSystem;
            SysDataToMerge.Remove(ent);
        }
    }
}
