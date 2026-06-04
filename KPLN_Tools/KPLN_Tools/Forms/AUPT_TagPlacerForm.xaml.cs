using Autodesk.Revit.DB;
using KPLN_Tools.Forms.Models;
using System.Windows;

namespace KPLN_Tools.Forms
{
    public partial class AUPT_TagPlacerForm : Window
    {
        public AUPT_TagPlacerForm(Document doc)
        {
            InitializeComponent();

            DataContext = new AUPTTagPlacerVM(doc);
        }

        public AUPTTagPlacerVM AUPTTagPlacerViewModel => DataContext as AUPTTagPlacerVM;
    }
}
