using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;

namespace KPLN_CoordiantorAI.Forms
{
    public partial class ExternalModelDescriptionsWindow : Window
    {
        private const string TitleBlockDescriptionName = "Описание параметров основной надписи";
        private readonly ObservableCollection<ExternalModelDescriptionRow> _rows;

        public ExternalModelDescriptionsWindow(string titleBlockParametersDescription)
        {
            _rows = new ObservableCollection<ExternalModelDescriptionRow>
            {
                new ExternalModelDescriptionRow
                {
                    Name = TitleBlockDescriptionName,
                    Text = titleBlockParametersDescription ?? string.Empty
                }
            };

            InitializeComponent();
            DescriptionsDataGrid.ItemsSource = _rows;
        }

        public string TitleBlockParametersDescription { get; private set; }

        private void OnSaveClick(object sender, RoutedEventArgs e)
        {
            ExternalModelDescriptionRow row = _rows.FirstOrDefault();
            TitleBlockParametersDescription = row == null ? string.Empty : row.Text ?? string.Empty;
            DialogResult = true;
        }

        private void OnCancelClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        public sealed class ExternalModelDescriptionRow
        {
            public string Name { get; set; }

            public string Text { get; set; }
        }
    }
}