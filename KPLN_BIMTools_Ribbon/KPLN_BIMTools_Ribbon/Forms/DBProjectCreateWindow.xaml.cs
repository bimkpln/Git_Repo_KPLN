using KPLN_BIMTools_Ribbon.Forms.ViewModels;
using System.Windows;

namespace KPLN_BIMTools_Ribbon.Forms
{
    public partial class DBProjectCreateWindow : Window
    {
        public DBProjectCreateWindow()
        {
            InitializeComponent();
            DataContext = new DBProjectCreateViewModel(this);
        }
    }
}
