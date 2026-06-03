using Autodesk.Revit.DB;
using KPLN_Tools.Forms.Models;
using System.Windows;

namespace KPLN_Tools.Forms
{
    public partial class AutosaveForm : Window
    {
        public AutosaveForm()
        {
            InitializeComponent();

            DataContext = new AutoSaveVM();
        }

        public AutoSaveVM ASVModel => DataContext as AutoSaveVM;
    }
}
