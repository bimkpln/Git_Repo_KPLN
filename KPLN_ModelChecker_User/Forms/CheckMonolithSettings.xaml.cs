using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_User.ExternalCommands;


namespace KPLN_ModelChecker_User.Forms
{
    /// <summary>
    /// Логика взаимодействия для CheckMonolithSettings.xaml
    /// </summary>
    public partial class CheckMonolithSettings : Window
    {
        UIApplication _uiApp;
        Document _activeDoc;
        private List<string> _categories = new List<string>();

        public IReadOnlyList<string> SelectedCategories => _categories;

        public double Tolerance { get; private set; }


        public CheckMonolithSettings(UIApplication uiApp, Document activeDoc)
        {
            InitializeComponent();

            _uiApp = uiApp;
            _activeDoc = activeDoc;
#if (Revit2023 || Debug2023)
            FillCategories();
            IList<FamilyInstance> monolithClashPoints = CommandCheckMonolith.GetMonolithClashPoints(_activeDoc);
            btnClashInfo.IsEnabled = monolithClashPoints.Count != 0;
#endif
        }

        /// <summary>
        /// Заполняем список категорий.
        /// </summary>
        private void FillCategories()
        {
            var categories = new List<string>
        {
            "Стены",
            "Перекрытия",
            "Лестницы"
        };
            lstCategories.ItemsSource = categories;
        }

        /// <summary>
        /// XAML. Кнопка действия. Проверить текущие ClashPoint
        /// </summary>
        private void ButtonClashInfo_Click(object sender, RoutedEventArgs e)
        {
#if (Revit2023 || Debug2023)
            IList<FamilyInstance> monolithClashPoints = CommandCheckMonolith.GetMonolithClashPoints(_activeDoc);
            if (monolithClashPoints.Count != 0)
            {
                var win = new CheckMonolithInfo(_uiApp.ActiveUIDocument, monolithClashPoints, null)
                {
                    Topmost = true,
                    ShowInTaskbar = false
                };
                win.Show();
                this.Close();
            }
#endif
        }

        /// <summary>
        /// XAML. Кнопка действия. Выбрать эллементы
        /// </summary>
        private void ButtonNext_Click(object sender, RoutedEventArgs e)
        {
            var chosenCategories = lstCategories.SelectedItems
                                                .OfType<string>()
                                                .ToList();

            if (chosenCategories.Count == 0)
            {
                MessageBox.Show(
                    "Выберите хотя бы одну категорию для проверки.",
                    "Недостаточно данных",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            if (!double.TryParse(txtTolerance.Text.Replace(',', '.'), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double tolerance)
                || tolerance < 0 || tolerance > 20)
            {
                MessageBox.Show(
                    "Введите допустимую погрешность от 0 до 20 (например, 0.01 или 5).",
                    "Недопустимое значение",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            _categories = chosenCategories;
            Tolerance = tolerance;
            DialogResult = true;
        }
    }
}