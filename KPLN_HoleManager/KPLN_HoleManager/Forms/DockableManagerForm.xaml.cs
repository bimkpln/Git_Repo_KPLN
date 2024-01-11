using Autodesk.Revit.UI;
using KPLN_HoleManager.Forms.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace KPLN_HoleManager.Forms
{
    public partial class DockableManagerForm : Page, IDockablePaneProvider
    {
        public DockableManagerForm()
        {
            InitializeComponent();
            //btnApprove.DataContext = new ButtonViewModel(
            //    new BitmapImage(new Uri("pack://application:,,,/KPLN_HoleManager.Forms;component/Imagens/DockableManager/Approve.png")),
            //    "Одобрить элемент(ы)",
            //    "Пометить выбранные элементы как «Утвержденные»");

            //btnAddSubElement.DataContext = "KPLN_HoleManager.Imagens.DockableManager.AddSubElements.png";
            //btnApplySubElements.DataContext = new Source(ExtensibleOpeningManager.Common.Collections.ImageButton.ApplySubElements);
            //btnApplyWall.DataContext = new Source(ExtensibleOpeningManager.Common.Collections.ImageButton.ApplyWall);
            //btnApprove.DataContext = new Source(ExtensibleOpeningManager.Common.Collections.ImageButton.Approve);
            //btnGroup.DataContext = new Source(ExtensibleOpeningManager.Common.Collections.ImageButton.Group);
            //btnReject.DataContext = new Source(ExtensibleOpeningManager.Common.Collections.ImageButton.Reject);
            //btnReset.DataContext = new Source(ExtensibleOpeningManager.Common.Collections.ImageButton.Reset);
            //btnSetOffset.DataContext = new Source(ExtensibleOpeningManager.Common.Collections.ImageButton.SetOffset);
            //btnSetWall.DataContext = new Source(ExtensibleOpeningManager.Common.Collections.ImageButton.SetWall);
            //btnSwap.DataContext = new Source(ExtensibleOpeningManager.Common.Collections.ImageButton.Swap);
            //btnUngroup.DataContext = new Source(ExtensibleOpeningManager.Common.Collections.ImageButton.Ungroup);
            //btnUpdate.DataContext = new Source(ExtensibleOpeningManager.Common.Collections.ImageButton.Update);
            //btnFindSubelements.DataContext = new Source(ExtensibleOpeningManager.Common.Collections.ImageButton.FindSubelements);
        }

        public void SetupDockablePane(DockablePaneProviderData data)
        {
            data.FrameworkElement = this as FrameworkElement;
            data.InitialState = new DockablePaneState
            {
                DockPosition = DockPosition.Tabbed,
                TabBehind = DockablePanes.BuiltInDockablePanes.ProjectBrowser
            };
        }

        public void UpdateItemscontroll()
        {
        }
        private void OnItemDoubleClick(object sender, MouseButtonEventArgs args)
        {
            
        }
        private void OnBtnApprove(object sender, RoutedEventArgs e)
        {
            
        }
        private void OnBtnReject(object sender, RoutedEventArgs e)
        {
        }
        private void OnBtnSetOffset(object sender, RoutedEventArgs e)
        {
        }
        private void OnBtnGroup(object sender, RoutedEventArgs e)
        {
        }
        private void OnBtnUngroup(object sender, RoutedEventArgs e)
        {
        }
        private void OnBtnReset(object sender, RoutedEventArgs e)
        {
        }
        private void OnBtnUpdate(object sender, RoutedEventArgs e)
        {
        }
        private void OnBtnApplySubElements(object sender, RoutedEventArgs e)
        {
        }
        private void OnBtnApplyWall(object sender, RoutedEventArgs e)
        {
        }
        private void OnBtnAddSubElement(object sender, RoutedEventArgs e)
        {
        }
        private void OnBtnSetWall(object sender, RoutedEventArgs e)
        {
        }
        private void OnBtnSwap(object sender, RoutedEventArgs e)
        {
        }
        //До тех пор пока не появится плагин для КР
        /*
        private void OnBtnPlaceOnKR(object sender, RoutedEventArgs e)
        {

        }
        private void OnBtnPlaceOnAR(object sender, RoutedEventArgs e)
        {

        }
        */
        private void OnBtnPlaceOnMEP(object sender, RoutedEventArgs args)
        {
            
        }
        private void OnBtnLoop(object sender, RoutedEventArgs args)
        {
            
        }
        private void OnBtnLoopDeny(object sender, RoutedEventArgs e)
        {
        }
        private void OnBtnLoopApply(object sender, RoutedEventArgs e)
        {
        }
        private void OnBtnLoopNext(object sender, RoutedEventArgs e)
        {
        }
        private void OnBtnLoopSkip(object sender, RoutedEventArgs e)
        {
        }
        private void OnBtnPlaceOnSelected(object sender, RoutedEventArgs e)
        {
            
        }

        private void OnBtnAddComment(object sender, RoutedEventArgs e)
        {
            
        }

        private void OnSubDepartmentChanged(object sender, SelectionChangedEventArgs args)
        {
           
        }

        private void OnSubItemRemoveBtnClick(object sender, RoutedEventArgs e)
        {
        }

        private void OnBtnPlaceOnSelectedTask(object sender, RoutedEventArgs e)
        {
        }

        private void OnSubItemAddRemarkBtnClick(object sender, RoutedEventArgs e)
        {
            
        }

        private void OnBtnFindSubelements(object sender, RoutedEventArgs args)
        {
            
        }
}
}
