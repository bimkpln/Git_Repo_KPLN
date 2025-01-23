using Autodesk.Revit.UI;
using KPLN_HoleManager.Common;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace KPLN_HoleManager.Forms
{
    public partial class DockableManagerForm : Page, IDockablePaneProvider
    {

        private readonly DBWorkerService _dbWorkerService;
        string userFullName; // Имя пользователя
        string departmentName; // Название отдела

        public DockableManagerForm()
        {
            InitializeComponent();

            _dbWorkerService = new DBWorkerService();

            // Получаем данные
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

        // Растановка кнопок в зависимости от отдела
        private void AddDepartmentButtons()
        {
            // Общий стиль кнопок
            var buttonStyle = new Style(typeof(Button));
            buttonStyle.Setters.Add(new Setter(Button.HeightProperty, 30.0));
            buttonStyle.Setters.Add(new Setter(Button.HorizontalAlignmentProperty, HorizontalAlignment.Stretch));
            buttonStyle.Setters.Add(new Setter(Button.HorizontalContentAlignmentProperty, HorizontalAlignment.Left));
            buttonStyle.Setters.Add(new Setter(Button.PaddingProperty, new Thickness(10, 0.5, 0, 0)));
            buttonStyle.Setters.Add(new Setter(Button.BackgroundProperty, new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFE0FDFF"))));
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

            ActionButtonDepartment.Children.Add(button);
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
