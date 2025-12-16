using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ExtraFilter.ExternalEventHandler;
using KPLN_ExtraFilter.Forms.ViewModels;
using System.Collections.Generic;
using System.Windows;

namespace KPLN_ExtraFilter.Forms
{
    public partial class SetParamsByFrameForm : Window
    {
        private ExternalEvent _viewExtEv;
        private ViewActivatedHandler _viewHandler;

        private ExternalEvent _selExtEv;
        private SelectionChangedHandler _selHandler;

        public SetParamsByFrameForm(Document doc, IEnumerable<Element> userSelElems)
        {
            CurrentSetParamsByFrameVM = new SetParamsByFrameVM(this, doc, userSelElems);

            InitializeComponent();

            DataContext = CurrentSetParamsByFrameVM;
        }

        /// <summary>
        /// VM для окна
        /// </summary>
        public SetParamsByFrameVM CurrentSetParamsByFrameVM { get; set; }

        public void SetExternalEvent(ExternalEvent viewExtEv, ViewActivatedHandler viewHandler, ExternalEvent selExtEv, SelectionChangedHandler selHandler)
        {
            _viewExtEv = viewExtEv;
            
            _viewHandler = viewHandler;
            _viewHandler.CurrentSetParamsByFrameVM = CurrentSetParamsByFrameVM;


            _selExtEv = selExtEv;
            
            _selHandler = selHandler;
            _selHandler.CurrentSetParamsByFrameVM = CurrentSetParamsByFrameVM;
        }

        public void RaiseUpdateSelChanged() => _selExtEv?.Raise();

        public void RaiseUpdateViewChanged() => _viewExtEv?.Raise();
    }
}
