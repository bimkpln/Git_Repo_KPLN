using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.ComponentModel;
using KPLN_HoleManager.Common;
using System.Collections.Generic;
using System.Linq;


namespace KPLN_HoleManager.Forms
{
    // Передача данных в функции
    public class HoleSelectionViewModel
    {
        public string UserFullName { get; }
        public string DepartmentName { get; }
        public string ElementInfo { get; }

        public HoleSelectionViewModel(string userFullName, string departmentName, Element element) {}
    }

    public partial class DockableManagerForm : Page, IDockablePaneProvider
    {
        UIApplication _uiApp; // Активная Revit-сессия
        Document _doc; // Активный Revit-документ

        private readonly DBWorkerService _dbWorkerService; // БД
        string userFullName; // Имя пользователя
        string departmentName; // Название отдела

        public DockableManagerForm()
        {
            InitializeComponent();
 
            // Получаем данные из БД
            _dbWorkerService = new DBWorkerService();            
            userFullName = _dbWorkerService.UserFullName;
            departmentName = _dbWorkerService.DepartmentName;

            // Построение интерфейса
            AddDepartmentButtons();
            DataContext = new ButtonDataViewModel();
        }

        // Регистрация кнопок в Revit
        public void SetupDockablePane(DockablePaneProviderData data)
        {
            data.FrameworkElement = this as FrameworkElement;
            data.InitialState = new DockablePaneState
            {
                DockPosition = DockPosition.Tabbed,
                TabBehind = DockablePanes.BuiltInDockablePanes.ProjectBrowser
            };
        }

        // Получение Reviot-потока и активного документа
        public void SetUIApplication(UIApplication uiApp)
        {
            _uiApp = uiApp;
            _doc = uiApp?.ActiveUIDocument?.Document;
        }

        // Растановка кнопок в зависимости от отдела
        private void AddDepartmentButtons()
        {
            // Общий стиль кнопок
            var buttonStyle = new Style(typeof(Button));
            buttonStyle.Setters.Add(new Setter(Button.HeightProperty, 30.0));
            buttonStyle.Setters.Add(new Setter(Button.HorizontalAlignmentProperty, HorizontalAlignment.Stretch));
            buttonStyle.Setters.Add(new Setter(Button.HorizontalContentAlignmentProperty, HorizontalAlignment.Left));
            buttonStyle.Setters.Add(new Setter(Button.PaddingProperty, new Thickness(10, 0.5, 0, 0)));
            buttonStyle.Setters.Add(new Setter(Button.BackgroundProperty, new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#FFE0FDFF"))));
            buttonStyle.Setters.Add(new Setter(Button.BorderThicknessProperty, new Thickness(1)));

            // Добавляем кнопки в зависимости от departmentName
            if (departmentName == "АР" || departmentName == "КР")
            {
                AddButton("🔄  Обновить отверстия по заданиям", buttonStyle);
                AddButton("🔂  Создать отверстия по заданию", buttonStyle);
                AddButton("➡️  Расставить отверстия по выбранной стене", buttonStyle);
            }
            else if (departmentName == "ОВиК" || departmentName == "ВК" || departmentName == "ЭОМ" || departmentName == "СС")
            {
                AddButton("🔂  Создать отверстия по заданию", buttonStyle);
                AddButton("➡️  Расставить отверстия по выбранной стене", buttonStyle);
                AddButton("🔀  Расставить отверстия по пересечениям", buttonStyle);
            }
            else if (departmentName == "BIM")
            {
                AddButton("🔄  Обновить отверстия по заданиям", buttonStyle);
                AddButton("🔂  Создать отверстия по заданию", buttonStyle);
                AddButton("➡️  Расставить отверстия по выбранной стене", buttonStyle);
                AddButton("🔀  Расставить отверстия по пересечениям", buttonStyle);
            }

            TestField.Text = $"{userFullName} ({departmentName})";
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
            if (content.Contains("Обновить отверстия по заданиям"))
            {
                button.Click += UpdateHolesByTasks;
            }
            else if (content.Contains("Создать отверстия по заданию"))
            {
                button.Click += CreateHolesByTask;
            }
            else if (content.Contains("Расставить отверстия по выбранной стене"))
            {
                button.Click += PlaceHolesOnSelectedWall;
            }
            else if (content.Contains("Расставить отверстия по пересечениям"))
            {
                button.Click += PlaceHolesByIntersections;
            }

            ActionButtonDepartment.Children.Add(button);
        }










        // Обработчик для кнопки "Расставить отверстия по выбранной стене"
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

            var holeWindow = new sChoiseHole(userFullName, departmentName, element);
            holeWindow.ShowDialog();
        }





        // Обработчик для кнопки "Обновить отверстия по заданиям"
        private void UpdateHolesByTasks(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Обновление отверстий по заданиям выполнено.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        // Обработчик для кнопки "Создать отверстия по заданию"
        private void CreateHolesByTask(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Создание отверстий по заданию выполнено.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        // Обработчик для кнопки "Расставить отверстия по пересечениям"
        private void PlaceHolesByIntersections(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Отверстия успешно расставлены по пересечениям!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }













    // Передаём данные статусов в названия кнопок
    public class ButtonDataViewModel : INotifyPropertyChanged
    {
        private string _noneStatusButtonText = $"❓  Без статуса: {null}";
        private string _approvedButtonText = $"✔️  Утверждено: {null}";
        private string _warningButtonText = $"⚠️  Предупреждения: {null}";
        private string _errorButtonText = $"❌  Ошибки: {null}";

        public string NoneStatusButtonText
        {
            get => _noneStatusButtonText;
            set
            {
                _approvedButtonText = value;
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

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}