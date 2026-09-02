using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_MEPSpacing.Forms.ViewModels;
using System.Collections.Generic;
using System.Windows;

namespace KPLN_MEPSpacing.Forms
{
    public partial class MepSpacingForm : Window
    {
        public MepSpacingForm(UIApplication uiapp, IEnumerable<ElementId> selectedElementIds)
        {
            CurrentMepSpacingVM = new MepSpacingVM(this, uiapp, selectedElementIds);

            InitializeComponent();

            DataContext = CurrentMepSpacingVM;
        }

        /// <summary>
        /// VM для окна.
        /// </summary>
        public MepSpacingVM CurrentMepSpacingVM { get; }
    }
}
