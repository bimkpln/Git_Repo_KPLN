using Autodesk.Revit.DB;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UIFactory;
using KPLN_Tools.ExecutableCommand;
using KPLN_Tools.Forms.Models;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace KPLN_Tools.Forms
{
    public partial class OVVK_SystemManagerForm : Window
    {
        public OVVK_SystemManagerForm(Document doc, List<Element> elemsInModel)
        {
            CurrentViewModel = new OVVK_SystemManager_ViewModel(doc, elemsInModel);

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
            ElementSinglePick paramForm = SelectParameterFromRevit.CreateForm(CurrentViewModel.CurrentDoc, CurrentViewModel.ElementColl, StorageType.String);
            paramForm.ShowDialog();

            if (paramForm.SelectedElement != null)
                CurrentViewModel.ParameterName = paramForm.SelectedElement.Name;

            UpdateEnable();
        }

        private void BtnCreateViews_Click(object sender, RoutedEventArgs e)
        {
            // Анализ вида
            Autodesk.Revit.DB.View activeView = CurrentViewModel.CurrentDoc.ActiveView;
            if (activeView == null || activeView.ViewType != ViewType.ThreeD)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Скрипт нужно запускать при открытом 3D-виде, т.к. на основании его будут создаваиться аналоги",
                    "Ошибка",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);

                return;
            }

            ElementMultiPick elementMultiPick = new ElementMultiPick(CurrentViewModel
                .SystemSumParameters
                .Where(pName => !pName.Contains("ВНИМАНИЕ!!!"))
                .Select(pName => new KPLN_Library_Forms.Common.ElementEntity(pName)));
            
            if ((bool)elementMultiPick.ShowDialog())
            {
                KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandSystemManager_ViewCreator(
                    CurrentViewModel.CurrentDoc,
                    elementMultiPick.SelectedElements.Select(ent => ent.Name).ToArray(),
                    CurrentViewModel.ParameterName, 
                    CurrentViewModel.SysNameSeparator));

                Close();
            }
        }

        private void BtnSelectWarningsElems_Click(object sender, RoutedEventArgs e)
        {
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandShowElement(CurrentViewModel.WarningsElementColl));
            Close();
        }

        private void UpdateEnable()
        {
            BtnCreateViews.IsEnabled = CurrentViewModel.SystemSumParameters.Any();
            BtnSelectWarningsElems.IsEnabled = CurrentViewModel.WarningsElementColl.Any();
        }
    }
}
