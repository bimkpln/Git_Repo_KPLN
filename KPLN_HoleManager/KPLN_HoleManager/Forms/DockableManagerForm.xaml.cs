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
using System.Windows.Documents;
using System.Text;


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

            AddButton("🔄  Принудительное обновление данных отверстий", buttonStyle);
            AddButton("➡️  Создать задание на отверстие", buttonStyle);
        }

        // Обновление статусов отверстий
        public void UpdateStatusCounts()
        {
            if (_uiApp == null || _uiApp.ActiveUIDocument == null) return;

            Document doc = _uiApp.ActiveUIDocument.Document;
            List<ElementId> familyInstanceIds = _iDataProcessor.GetFamilyInstanceIds(doc);
            List<int> statusCounts = _iDataProcessor.StatusHoleTask(doc, familyInstanceIds);

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
            if (content.Contains("Принудительное обновление данных отверстий"))
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

        // XAML. Вывод информации об отверстиях
        private void StatusButton_Click(object sender, RoutedEventArgs e)
        {
            Button clickedButton = sender as Button;
            if (clickedButton == null) return;

            // Очищаем панель перед добавлением новых кнопок
            InfoHolePanel.Children.Clear();

            // Получаем список всех сообщений для отверстий
            Document doc = _uiApp.ActiveUIDocument.Document;
            List<ElementId> familyInstanceIds = _iDataProcessor.GetFamilyInstanceIds(doc);
            List<List<string>> holeTaskMessages = _iDataProcessor.GetHoleLastTaskMessages(doc, familyInstanceIds);

            // Определяем нужный статус на основе нажатой кнопки
            string selectedStatus = null;
            if (clickedButton.Name == "ButtonNoneStatus") selectedStatus = "Без статуса";
            else if (clickedButton.Name == "ButtonOKStatus") selectedStatus = "Утверждено";
            else if (clickedButton.Name == "ButtonWarningStatus") selectedStatus = "Предупреждения";
            else if (clickedButton.Name == "ButtonErrorStatus") selectedStatus = "Ошибки";            

            // Фильтруем сообщения по выбранному статусу
            List<List<string>> filteredMessages = new List<List<string>>();

            foreach (var messageParts in holeTaskMessages)
            {
                bool matchesDepartment = departmentName == "BIM" ||
                                         messageParts[2] == departmentName ||
                                         messageParts[3] == departmentName;

                if (matchesDepartment && messageParts[9] == selectedStatus)
                {
                    filteredMessages.Add(messageParts);
                }
            }

            foreach (var messageParts in filteredMessages)
            {
                // Разбиваем messageParts на отдельные переменные
                string data = messageParts[0];
                string name = messageParts[1];
                string departamentFrom = messageParts[2];
                string departamentIn = messageParts[3];
                string holeID = messageParts[4];
                string holeName = messageParts[5];
                string wallID = messageParts[6];
                string sEllementID = messageParts[7];

                // Создаем TextBlock для кнопки
                TextBlock textBlock = new TextBlock
                {
                    TextWrapping = TextWrapping.Wrap
                };

                textBlock.Inlines.Add(new Run(data) { FontWeight = FontWeights.Bold, Foreground = Brushes.Blue });
                textBlock.Inlines.Add(new Run($". {name} ({departamentFrom} -> {departamentIn}).\n"));
                textBlock.Inlines.Add(new Run(holeName) { FontWeight = FontWeights.Bold });
                textBlock.Inlines.Add(new Run($" ({holeID}).\n"));
                textBlock.Inlines.Add(new Run("Стена: ") { FontWeight = FontWeights.Bold });
                textBlock.Inlines.Add(new Run($"{wallID}. "));
                textBlock.Inlines.Add(new Run("Элементы в стене: ") { FontWeight = FontWeights.Bold });
                textBlock.Inlines.Add(new Run($"{sEllementID}."));

                Button newButton = new Button
                {
                    Content = textBlock,
                    Height = 68,
                    HorizontalAlignment = HorizontalAlignment.Stretch,
                    HorizontalContentAlignment = HorizontalAlignment.Left, // Выравнивание текста влево
                    Margin = new Thickness(2, 2, 2, 0), // Отступы между кнопками
                    Padding = new Thickness(10, 0, 0, 0) // Внутренний отступ кнопки
                };

                // Определяем цвет кнопки на основе статуса
                if (selectedStatus == "Без статуса")
                    newButton.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(235, 235, 235)); // Серый
                else if (selectedStatus == "Утверждено")
                    newButton.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(188, 252, 157)); // Зеленый
                else if (selectedStatus == "Предупреждения")
                    newButton.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(253, 245, 138)); // Желтый
                else if (selectedStatus == "Ошибки")
                    newButton.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 192, 177)); // Красный

                // Добавляем обработчик нажатия
                newButton.Click += (s, ev) =>
                {
                    InfoHolePanel.Children.Clear(); // Очищаем панель

                    // Базовая информация**
                    TextBlock generalInfoTextBlock = new TextBlock { TextWrapping = TextWrapping.Wrap, Padding = new Thickness(8) };
                    generalInfoTextBlock.Inlines.Add(new Run(holeName) { FontWeight = FontWeights.Bold });
                    generalInfoTextBlock.Inlines.Add(new Run($" ({holeID})\n"));

                    Brush statusColor = Brushes.Gray;
                    string statusText = messageParts[9];

                    switch (statusText)
                    {
                        case "Без статуса":
                            statusColor = Brushes.Gray;
                            break;
                        case "Утверждено":
                            statusColor = Brushes.Green;
                            break;
                        case "Предупреждения":
                            statusColor = Brushes.DarkGoldenrod;
                            break;
                        case "Ошибки":
                            statusColor = Brushes.Red;
                            break;
                    }

                    generalInfoTextBlock.Inlines.Add(new Run(statusText) { FontWeight = FontWeights.Bold, Foreground = statusColor });
                    generalInfoTextBlock.Inlines.Add(new Run($"\n"));
                    generalInfoTextBlock.Inlines.Add(new Run("Стена: ") { FontWeight = FontWeights.Bold });
                    generalInfoTextBlock.Inlines.Add(new Run($"{wallID}\n"));
                    generalInfoTextBlock.Inlines.Add(new Run("Элементы в отверстии: ") { FontWeight = FontWeights.Bold });
                    generalInfoTextBlock.Inlines.Add(new Run($"{sEllementID}"));

                    // Получаем список сообщений
                    List<List<string>> fullHoleInfo = _iDataProcessor.GetHoleTaskMessages(doc, holeID);

                    // Контейнер для сообщений
                    StackPanel messagesPanel = new StackPanel { Margin = new Thickness(0, 5, 0, 0) };

                    foreach (List<string> fullHoleInfoParts in fullHoleInfo)
                    {
                        if (fullHoleInfoParts.Count > 10) // Проверяем, что в списке достаточно элементов
                        {
                            // Извлекаем нужные данные
                            string mDate = fullHoleInfoParts[0];
                            string mName = fullHoleInfoParts[1];
                            string mDepartmentFrom = fullHoleInfoParts[2];
                            string mDepartmentTo = fullHoleInfoParts[3];
                            string mCoordinates = fullHoleInfoParts[8];
                            string mMessageText = fullHoleInfoParts[10];

                            // Создаем TextBlock для каждого сообщения
                            TextBlock messageTextBlock = new TextBlock { TextWrapping = TextWrapping.Wrap };

                            messageTextBlock.Inlines.Add(new Run($"{mDate}") { FontWeight = FontWeights.Bold, Foreground = Brushes.BlueViolet });
                            messageTextBlock.Inlines.Add(new Run($" | {mName} ({mDepartmentFrom} → {mDepartmentTo})\n"));

                            messageTextBlock.Inlines.Add(new Run("Координаты: ") { FontWeight = FontWeights.Bold });
                            messageTextBlock.Inlines.Add(new Run($"{mCoordinates}\n"));

                            messageTextBlock.Inlines.Add(new Run("💬  Сообщение: ") { FontWeight = FontWeights.Bold });
                            messageTextBlock.Inlines.Add(new Run($"{mMessageText}"));

                            Border messageBorder = new Border
                            {
                                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(230, 245, 255)),
                                BorderBrush = Brushes.LightGray,
                                BorderThickness = new Thickness(1),
                                CornerRadius = new CornerRadius(5),
                                Padding = new Thickness(8),
                                Child = messageTextBlock
                            };

                            messagesPanel.Children.Add(messageBorder);
                        }
                    }

                    // Если сообщений нет
                    if (messagesPanel.Children.Count == 0)
                    {
                        messagesPanel.Children.Add(new TextBlock
                        {
                            Text = "Нет информации об отверстии",
                            TextWrapping = TextWrapping.Wrap,
                            Foreground = Brushes.Gray
                        });
                    }

                    // Основной контейнер
                    StackPanel contentPanel = new StackPanel { Margin = new Thickness(5, 5, 5, 0) };
                    contentPanel.Children.Add(generalInfoTextBlock);
                    contentPanel.Children.Add(messagesPanel);

                    // Добавляем в основную панель
                    InfoHolePanel.Children.Add(contentPanel);





                    // Перемещение к элементу с ID holeID на 3D виде
                    if (int.TryParse(holeID, out int elementIdValue))
                    {
                        ElementId elementId = new ElementId(elementIdValue);
                        Element element = doc.GetElement(elementId);

                        if (element != null)
                        {
                            UIDocument uiDoc = _uiApp.ActiveUIDocument;

                            // Проверяем, есть ли открытый 3D-вид
                            View3D active3DView = null;
                            View currentView = doc.ActiveView;

                            if (currentView is View3D view3D && !view3D.IsTemplate)
                            {
                                active3DView = view3D; 
                            }
                            else
                            {

                                active3DView = new FilteredElementCollector(doc)
                                    .OfClass(typeof(View3D))
                                    .Cast<View3D>()
                                    .FirstOrDefault(v => !v.IsTemplate);
                            }

                            if (active3DView != null)
                            {
                                if (doc.ActiveView.Id != active3DView.Id)
                                {
                                    uiDoc.ActiveView = active3DView;
                                }

                                uiDoc.Selection.SetElementIds(new List<ElementId> { elementId });

                                uiDoc.ShowElements(element);
                            }
                        }
                    }
                };

                // Добавляем кнопку в панель
                InfoHolePanel.Children.Add(newButton);
            }
        }
    }

    // Передаём данные статусов в названия кнопок
    public class ButtonDataViewModel : INotifyPropertyChanged
    {
        // Первичная прогрузка
        private string _noneStatusButtonText = "❓  Без статуса: ОШИБКА!";
        private string _approvedButtonText = "✔️  Утверждено: ОШИБКА!";
        private string _warningButtonText = "⚠️  Предупреждения: ОШИБКА!";
        private string _errorButtonText = "❌  Ошибки: ОШИБКА!";

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