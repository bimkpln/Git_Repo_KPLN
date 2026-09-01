using Autodesk.Revit.UI;
using KPLN_TrailingMEP.Forms.ViewModels;
using System.Windows;

namespace KPLN_TrailingMEP.Forms
{
    public partial class RouteTraceForm : Window
    {
        public RouteTraceForm(UIApplication uiapp)
        {
            CurrentRouteTraceVM = new RouteTraceVM(this, uiapp);

            InitializeComponent();

            DataContext = CurrentRouteTraceVM;
        }

        /// <summary>
        /// VM для окна
        /// </summary>
        public RouteTraceVM CurrentRouteTraceVM { get; set; }
    }
}
