using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.ComponentModel;
using KPLN_HoleManager.Common;
using System.Collections.Generic;
using System.Linq;
using KPLN_HoleManager.Commands;


namespace KPLN_HoleManager.Forms
{
    // Передача данных в функции
    public class HoleSelectionViewModel
    {
        public string UserFullName { get; }
        public string DepartmentName { get; }

        public HoleSelectionViewModel(Element element, string userFullName, string departmentName) {}
    }

    public partial class DockableManagerForm : Page, IDockablePaneProvider
    {
        UIApplication _uiApp; // Активная Revit-сессия      

        private readonly DBWorkerService _dbWorkerService; // БД
        string userFullName; // Имя пользователя
        string departmentName; // Название отдела

        // Данные статусов в названия кнопок
        private readonly ButtonDataViewModel _buttonDataViewModel; 
        private static DockableManagerForm _instance;
        public static DockableManagerForm Instance => _instance;

        /// Получение Revit-потока
        public void SetUIApplication(UIApplication uiApp)
        {
            _uiApp = uiApp;
            UpdateStatusCounts();
        }

        /// Регистрация Dockable-панели в Revit
        public void SetupDockablePane(DockablePaneProviderData data)
        {
            data.FrameworkElement = this as FrameworkElement;
            data.InitialState = new DockablePaneState
            {
                DockPosition = DockPosition.Tabbed,
                TabBehind = DockablePanes.BuiltInDockablePanes.ProjectBrowser
            };
        }

        /// DockableManagerForm
        public DockableManagerForm()
        {
            InitializeComponent();
 
            // Получаем данные из БД
            _dbWorkerService = new DBWorkerService();            
            userFullName = _dbWorkerService.UserFullName;
            departmentName = _dbWorkerService.DepartmentName;

            // Кнопки характерные для отдела
            AddDepartmentButtons();

            // Общие кнопки для всех
            _buttonDataViewModel = new ButtonDataViewModel();
            DataContext = _buttonDataViewModel;
            _instance = this;
        }

        // Растановка кнопок в зависимости от отдела
        public void AddDepartmentButtons()
        {
            // Общий стиль кнопок
            var buttonStyle = new Style(typeof(Button));
            buttonStyle.Setters.Add(new Setter(Button.HeightProperty, 30.0));
            buttonStyle.Setters.Add(new Setter(Button.HorizontalAlignmentProperty, HorizontalAlignment.Stretch));
            buttonStyle.Setters.Add(new Setter(Button.HorizontalContentAlignmentProperty, HorizontalAlignment.Left));
            buttonStyle.Setters.Add(new Setter(Button.PaddingProperty, new Thickness(10, 0.5, 0, 0)));
            buttonStyle.Setters.Add(new Setter(Button.BackgroundProperty, new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#FFE0FDFF"))));
            buttonStyle.Setters.Add(new Setter(Button.BorderThicknessProperty, new Thickness(1)));

            AddButton("🔄  Обновление данных", buttonStyle);
            AddButton("➡️  Создать задание на отверстие", buttonStyle);
        }

        // Обновление статусов отверстий
        public void UpdateStatusCounts()
        {
            if (_uiApp == null || _uiApp.ActiveUIDocument == null) return;

            Document doc = _uiApp.ActiveUIDocument.Document;
            List<ElementId> familyInstanceIds = _iDataProcessor.GetFamilyInstanceIds(doc);
            List<int> statusCounts = _iDataProcessor.statusHoleTask(doc, familyInstanceIds);

            _buttonDataViewModel.UpdateStatusCounts(statusCounts);
        }

        // Функция пакетного создания кнопок
        private void AddButton(string content, Style style)
        {
            var button = new Button
            {
                Content = content,
                Style = style
            };

            // Добавляем обработчики для кнопок в зависимости от их содержимого
            if (content.Contains("Обновление данных"))
            {
                button.Click += UpdateHoles;
            }
            if (content.Contains("Создать задание на отверстие"))
            {
                button.Click += PlaceHolesOnSelectedWall;
            }

            ActionButtonDepartment.Children.Add(button);
        }

        // XAML. Обработчик для кнопки "Обновление данных об отверстиях"
        private void UpdateHoles(object sender, RoutedEventArgs e)
        {
            UpdateStatusCounts();
            TaskDialog.Show("Обновление данных", "Данные об отверстиях обновлены.");
        }

        // XAML. Обработчик для кнопки "Создать задание на отверстие"
        private void PlaceHolesOnSelectedWall(object sender, RoutedEventArgs e)
        {
            UIDocument uiDoc = _uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            ICollection<ElementId> selectedIds = uiDoc.Selection.GetElementIds();

            if (selectedIds.Count == 0)
            {
                TaskDialog.Show("Предупреждение", "Ничего не выбрано.\nПожалуйста, выберите стену.");
                return;
            }
            else if (selectedIds.Count > 1)
            {
                TaskDialog.Show("Предупреждение", "Выбрано несколько элементов.\nПожалуйста, выберите только один элемент.");
                return;
            }

            ElementId selectedId = selectedIds.First();
            Element element = doc.GetElement(selectedId);

            if (!(element is Wall wall))
            {
                TaskDialog.Show("Предупреждение", "Выбранный элемент не является стеной.\nПожалуйста, выберите стену.");
                return;
            }

            var holeWindow = new sChoiseHole(_uiApp, element, userFullName, departmentName);
            holeWindow.ShowDialog();
        }
    }

    // Передаём данные статусов в названия кнопок
    public class ButtonDataViewModel : INotifyPropertyChanged
    {
        // Первичная прогрузка
        private string _noneStatusButtonText = "❓  Без статуса: ?";
        private string _approvedButtonText = "✔️  Утверждено: ?";
        private string _warningButtonText = "⚠️  Предупреждения: ?";
        private string _errorButtonText = "❌  Ошибки: ?";

        public string NoneStatusButtonText
        {
            get => _noneStatusButtonText;
            set
            {
                _noneStatusButtonText = value;
                OnPropertyChanged(nameof(NoneStatusButtonText));
            }
        }

        public string ApprovedButtonText
        {
            get => _approvedButtonText;
            set
            {
                _approvedButtonText = value;
                OnPropertyChanged(nameof(ApprovedButtonText));
            }
        }

        public string WarningButtonText
        {
            get => _warningButtonText;
            set
            {
                _warningButtonText = value;
                OnPropertyChanged(nameof(WarningButtonText));
            }
        }

        public string ErrorButtonText
        {
            get => _errorButtonText;
            set
            {
                _errorButtonText = value;
                OnPropertyChanged(nameof(ErrorButtonText));
            }
        }

        // Обновление статусов отверстий в интерфесе
        public void UpdateStatusCounts(List<int> statusCounts)
        {
            if (statusCounts == null) return;

            NoneStatusButtonText = $"❓  Без статуса: {statusCounts[0]}";
            ApprovedButtonText = $"✔️  Утверждено: {statusCounts[1]}";
            WarningButtonText = $"⚠️  Предупреждения: {statusCounts[2]}";
            ErrorButtonText = $"❌  Ошибки: {statusCounts[3]}";
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}