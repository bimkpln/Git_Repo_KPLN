using System.Windows;

namespace KPLN_Tools.Forms
{
    public partial class NodeManagerInputTagWindow : Window
    {
        public string TagText { get; private set; }

        public NodeManagerInputTagWindow()
        {
            InitializeComponent();

            Loaded += (s, e) =>
            {
                TxtTag.Focus();
                TxtTag.SelectAll();
            };
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            TagText = TxtTag.Text?.Trim();

            if (string.IsNullOrWhiteSpace(TagText))
            {
                MessageBox.Show(this, "Тег не может быть пустым.", "Добавить тег", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            DialogResult = true;
            Close();
        }
    }
}
