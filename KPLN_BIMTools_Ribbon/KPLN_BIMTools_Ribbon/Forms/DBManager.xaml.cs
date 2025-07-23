using KPLN_BIMTools_Ribbon.Forms.ViewModels;
using System.Windows;
using System.Windows.Input;

namespace KPLN_BIMTools_Ribbon.Forms
{
    public partial class DBManager : Window
    {

        public DBManager()
        {
            InitializeComponent();

            DataContext = new DBManagerViewModel();

            PreviewKeyDown += new KeyEventHandler(HandleEsc);
        }

        private void HandleEsc(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
        }
    }
}
