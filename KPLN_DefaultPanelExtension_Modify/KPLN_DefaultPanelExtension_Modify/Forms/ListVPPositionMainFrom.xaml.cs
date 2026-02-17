using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_DefaultPanelExtension_Modify.ExternalEventHandler;
using KPLN_DefaultPanelExtension_Modify.Forms.Models;
using System.Windows;
using System.Windows.Input;

namespace KPLN_DefaultPanelExtension_Modify.Forms
{
    public partial class ListVPPositionMainFrom : Window
    {
        private ExternalEvent _extEv;
        private ListVPPositionHandler _handler;

        public ListVPPositionMainFrom(Element[] selTrueElems)
        {
            CurrentListVPPositionMainVM = new ListVPPositionMainVM(this, selTrueElems);
            
            InitializeComponent();

            DataContext = CurrentListVPPositionMainVM;
        }

        /// <summary>
        /// VM для окна
        /// </summary>
        public ListVPPositionMainVM CurrentListVPPositionMainVM { get; set; }

        public void SetExternalEvent(ExternalEvent extEv, ListVPPositionHandler handler)
        {
            _extEv = extEv;

            _handler = handler;
            _handler.HandlerListVPPositionMainVM = CurrentListVPPositionMainVM;
        }

        public void RaiseSelChanged() => _extEv?.Raise();

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (DataContext is ListVPPositionMainVM vm)
                vm.SaveConfig();
        }
    }
}
