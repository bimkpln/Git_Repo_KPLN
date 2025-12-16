using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ExtraFilter.Forms.ViewModels;
using System.Windows;

namespace KPLN_ExtraFilter.Forms
{
    public partial class SearchByIdForm : Window
    {
        public SearchByIdForm(UIApplication uiapp, View3D special3DView)
        {
            CurrentSearchByIdVM = new SearchByIdVM(this, uiapp, special3DView);

            InitializeComponent();

            this.tbUserInput.Focus();
            DataContext = CurrentSearchByIdVM;
        }

        /// <summary>
        /// VM для окна
        /// </summary>
        public SearchByIdVM CurrentSearchByIdVM { get; set; }
    }
}
