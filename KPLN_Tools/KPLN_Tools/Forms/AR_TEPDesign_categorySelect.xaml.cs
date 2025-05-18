using System.Windows;
using System.Windows.Controls;

namespace KPLN_Tools.Forms
{
    /// <summary>
    /// Логика взаимодействия для AR_TEPDesign_categorySelect.xaml
    /// </summary>
    public partial class AR_TEPDesign_categorySelect : Window
    {
        public int Result { get; private set; } = 0;

        public AR_TEPDesign_categorySelect()
        {
            InitializeComponent();
        }

        private void CategoryButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && int.TryParse(button.Tag?.ToString(), out int value))
            {
                Result = value;
                Close();
            }
        }
    }
}
