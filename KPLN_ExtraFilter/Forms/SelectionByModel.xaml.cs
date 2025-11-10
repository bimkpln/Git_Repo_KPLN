using System.Windows;
using System.Windows.Input;

namespace KPLN_ExtraFilter.Forms
{
    public partial class SelectionByModel : Window
    {

        public SelectionByModel()
        {
            InitializeComponent();



            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
        }
    }
}
