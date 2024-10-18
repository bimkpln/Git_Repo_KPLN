using Autodesk.Revit.DB;
using KPLN_Tools.Common.OVVK_System;
using KPLN_Tools.ExecutableCommand;
using KPLN_Tools.Forms.Models;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;

namespace KPLN_Tools.Forms
{
    public partial class OVVK_SystemManagerForm : Window
    {
        private readonly List<Element> _elemsInModel;

        public OVVK_SystemManagerForm(List<Element> elemsInModel)
        {
            _elemsInModel = elemsInModel;
            CurrentViewModel = new OVVK_SystemManager_ViewModel()
            {
                ParameterName = "ASML_Имя системы",
                SysNameSeparator = "/",
                SystemSumParameters = new ObservableCollection<string>() { "<данный функционал в разработке...>", "<данный функционал в разработке...>" }
            };

            InitializeComponent();

            this.DataContext = CurrentViewModel;
            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        public OVVK_SystemManager_ViewModel CurrentViewModel { get; private set; }

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
        }

        private void BtnParamSearch_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("упс.... в разработке, пока пиши имя вручную");
        }

        private void BtnCreateViews_Click(object sender, RoutedEventArgs e)
        {
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandSystemManager_ViewCreator(CurrentViewModel.ParameterName, CurrentViewModel.SysNameSeparator));

            Close();
        }
    }
}
