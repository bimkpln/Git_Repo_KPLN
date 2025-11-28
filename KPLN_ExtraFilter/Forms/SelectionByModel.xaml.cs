using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ExtraFilter.ExternalEventHandler;
using KPLN_ExtraFilter.Forms.ViewModels;
using System.Windows;

namespace KPLN_ExtraFilter.Forms
{
    public partial class SelectionByModel : Window
    {
        private ExternalEvent _viewExtEv;
        private ViewActivatedHandler _viewHandler;

        private ExternalEvent _selExtEv;
        private SelectionChangedHandler _selHandler;

        public SelectionByModel(Document doc)
        {
            CurrentSelectionByModelVM = new SelectionByModelVM(this, doc);

            InitializeComponent();

            DataContext = CurrentSelectionByModelVM;

#if Debug2020 || Revit2020
            // Нет метода в API для отслеживания изменний в выборке юзера
            this.AddToCurrent.Visibility = System.Windows.Visibility.Hidden;
#endif
        }

        /// <summary>
        /// VM для окна
        /// </summary>
        public SelectionByModelVM CurrentSelectionByModelVM { get; set; }

        public void SetExternalEvent(ExternalEvent viewExtEv, ViewActivatedHandler viewHandler, ExternalEvent selExtEv, SelectionChangedHandler selHandler)
        {
            _viewExtEv = viewExtEv;
            _viewHandler = viewHandler;
            _viewHandler.CurrentSelByModelVM = CurrentSelectionByModelVM;

            _selExtEv = selExtEv;
            _selHandler = selHandler;
            _selHandler.CurrentSelByModelVM = CurrentSelectionByModelVM;
        }

        public void RaiseUpdateSelChanged() => _selExtEv?.Raise();

        public void RaiseUpdateViewChanged() => _viewExtEv?.Raise();

        private void CHB_Where_Workset_Checked(object sender, RoutedEventArgs e) => this.CB_FilterWS.Focus();
    }
}
