using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;


namespace KPLN_ModelChecker_User.Forms
{
    /// <summary>
    /// Логика взаимодействия для CheckMonolithSettings.xaml
    /// </summary>
    public partial class CheckMonolithSettings : Window
    {
        Document _activeDoc; 
        private readonly List<Document> _allModels = new List<Document>();

        private Document _mainDoc;                 
        private List<Document> _linkedDocs = new List<Document>();
        private List<string> _categories = new List<string>();

        public Document MainDocument => _mainDoc;
        public IReadOnlyList<Document> LinkedDocuments => _linkedDocs;
        public IReadOnlyList<string> SelectedCategories => _categories;

        public double Tolerance { get; private set; }

        public CheckMonolithSettings(UIApplication uiApp, Document activeDoc)
        {
            InitializeComponent();

            _activeDoc = activeDoc;
            CollectModels(activeDoc);

            BindModels();
            FillCategories();
        }

        /// <summary>
        /// Собираем проект и все подгруженные в него Revit-связи.
        /// </summary>
        private void CollectModels(Document activeDoc)
        {
            _allModels.Clear();
            _allModels.Add(activeDoc);     

            var linkInstances = new FilteredElementCollector(activeDoc)
                                .OfClass(typeof(RevitLinkInstance))
                                .Cast<RevitLinkInstance>();

            foreach (RevitLinkInstance link in linkInstances)
            {
                Document linkDoc = link.GetLinkDocument();
                if (linkDoc != null && !_allModels.Contains(linkDoc))
                    _allModels.Add(linkDoc);
            }
        }

        /// <summary>
        /// Заполняем ComboBox и ListBox, подписываемся на изменения.
        /// </summary>
        private void BindModels()
        {
            cmbMainModel.ItemsSource = _allModels;
            cmbMainModel.DisplayMemberPath = "Title";
            cmbMainModel.SelectedItem = _activeDoc;
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
        /// XAML. Каждый раз, когда пользователь меняет «основную» модель, пересчитываем «связные»
        /// </summary>
        private void cmbMainModel_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Document mainDoc = cmbMainModel.SelectedItem as Document;
            if (mainDoc == null)
                return;

            var rest = _allModels.Where(d => d != mainDoc).ToList();

            lstLinkedModels.ItemsSource = rest;
            lstLinkedModels.DisplayMemberPath = "Title";
        }

        /// <summary>
        /// XAML. Кнопка действия
        /// </summary>
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Document mainDoc = cmbMainModel.SelectedItem as Document;
            if (mainDoc == null)
            {
                MessageBox.Show("Выберите основную модель", "Недостаточно данных",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var linkedDocs = lstLinkedModels.SelectedItems
                                            .OfType<Document>()
                                            .ToList();

            if (linkedDocs.Count == 0)
            {
                MessageBox.Show(
                    "Выберите хотя бы одну связную модель.",
                    "Недостаточно данных",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

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

            _mainDoc = mainDoc;
            _linkedDocs = linkedDocs;
            _categories = chosenCategories;
            Tolerance = tolerance;
            DialogResult = true;
        }
    }
}
