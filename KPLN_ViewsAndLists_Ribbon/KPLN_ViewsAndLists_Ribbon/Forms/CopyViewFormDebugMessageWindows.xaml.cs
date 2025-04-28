using System.Windows;

namespace KPLN_ViewsAndLists_Ribbon.Forms
{
    public partial class DebugMessageWindow : Window
    {
        public DebugMessageWindow(string message)
        {
            InitializeComponent();
            MessageTextBox.Text = message;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}