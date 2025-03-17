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
using System;
using Autodesk.Revit.UI.Selection;


namespace KPLN_HoleManager.Forms
{
    // Передача данных в функции
    public class HoleSelectionViewModel
    {
        public HoleSelectionViewModel(Element element, bool wallLink, string userFullName, string departmentName) { }
        public string UserFullName { get; }
        public string DepartmentName { get; }   
    }

    public partial class DockableManagerForm : Page, IDockablePaneProvider
    {
        UIApplication _uiApp;     

        private readonly DBWorkerService _dbWorkerService; 
        string userFullName; 
        string departmentName; 

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

        // Обновление статусов отверстий
        public void UpdateStatusCounts()
        {
            if (_uiApp == null || _uiApp.ActiveUIDocument == null) return;

            Document doc = _uiApp.ActiveUIDocument.Document;
            List<ElementId> familyInstanceIds = _iDataProcessor.GetFamilyInstanceIds(doc);
            List<int> statusCounts = _iDataProcessor.StatusHoleTask(doc, familyInstanceIds, userFullName, departmentName);

            _buttonDataViewModel.UpdateStatusCounts(statusCounts);
        }

        // Растановка кнопок в зависимости от отдела
        public void AddDepartmentButtons()
        {
            var buttonStyle = new Style(typeof(Button));
            buttonStyle.Setters.Add(new Setter(Button.HeightProperty, 30.0));
            buttonStyle.Setters.Add(new Setter(Button.HorizontalAlignmentProperty, HorizontalAlignment.Stretch));
            buttonStyle.Setters.Add(new Setter(Button.HorizontalContentAlignmentProperty, HorizontalAlignment.Left));
            buttonStyle.Setters.Add(new Setter(Button.PaddingProperty, new Thickness(10, 0.5, 0, 0)));
            buttonStyle.Setters.Add(new Setter(Button.BorderThicknessProperty, new Thickness(1)));

            AddButton("🔄  Обновление данных отверстий", buttonStyle, "#d1f7ff"); 
            AddButton("➡️  Создать задание на отверстие", buttonStyle, "#d1f7ff");
            AddButton("➡️  Создать задание по стене", buttonStyle, "#d1f7ff");
            AddButton("⚙  Настройка плагина", buttonStyle, "#d1f7ff");
        }

        // Пакетного создания кнопок
        public void AddButton(string content, Style baseStyle, string backgroundColor)
        {
            var button = new Button
            {
                Content = content,
                Style = baseStyle,
                Background = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString(backgroundColor))
            };

            // Добавляем обработчики для кнопок в зависимости от их содержимого
            if (content.Contains("Обновление данных отверстий"))
            {
                button.Click += UpdateHoles;
            }
            if (content.Contains("Создать задание на отверстие"))
            {
                button.Click += PlaceHolesOnSelectedWall;
            }
            if (content.Contains("Создать задание по стене"))
            {                
            }
            if (content.Contains("Настройка плагина"))
            {
                button.Click += HolePluginSettings;
            }

            ActionButtonDepartment.Children.Add(button);
        }

       // XAML. Обработчик для кнопки "Обновление данных об отверстиях"
        public void UpdateHoles(object sender, RoutedEventArgs e)
        {
            InfoHolePanel.Children.Clear();
            InfoHolePanel.RowDefinitions.Clear();
            UpdateStatusCounts();
            TaskDialog.Show("Обновление данных", "Данные об отверстиях обновлены.");
        }









        // XAML. Обработчик для кнопки "Создать задание на отверстие"
        public void PlaceHolesOnSelectedWall(object sender, RoutedEventArgs e)
        {
            UIDocument uiDoc = _uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;
            bool wallLink = false;

            InfoHolePanel.Children.Clear();
            InfoHolePanel.RowDefinitions.Clear();

            try
            {
                // Отключаем UI
                this.IsEnabled = false;

                // 🟢 ШАГ 1: Выбираем стену или ссылку
                Reference pickedRef = uiDoc.Selection.PickObject(
                    ObjectType.Element,
                    new WallAndLinkSelectionFilter(),
                    "Выберите стену или связанный элемент"
                );

                Element selectedElement = doc.GetElement(pickedRef.ElementId);

                // Если выбрана обычная стена
                if (selectedElement is Wall wall)
                {
                    wallLink = false;
                    ProcessHolePlacement(uiDoc, wall, wallLink);
                    return;
                }
                // Если выбран RevitLinkInstance
                else if (selectedElement is RevitLinkInstance linkInstance)
                {
                    Document linkedDoc = linkInstance.GetLinkDocument();

                    if (linkedDoc == null)
                    {
                        TaskDialog.Show("Ошибка", "Не удалось получить связанный документ.");
                        this.IsEnabled = true;
                        return;
                    }

                    // 🟢 ШАГ 2: Выбираем элемент внутри линка
                    Reference linkedRef = uiDoc.Selection.PickObject(
                        ObjectType.LinkedElement,
                        "Выберите стену в линке"
                    );

                    Element linkedElement = linkedDoc.GetElement(linkedRef.LinkedElementId);

                    if (linkedElement is Wall linkedWall)
                    {
                        wallLink = true;
                        ProcessHolePlacement(uiDoc, linkedWall, wallLink);
                        return;
                    }
                }
                else
                {
                    TaskDialog.Show("Ошибка", "Выбранный элемент не является стеной.");
                    this.IsEnabled = true;
                    return;
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                TaskDialog.Show("Отмена", "Выбор отменён пользователем.");
                this.IsEnabled = true;
                return;
            }

            TaskDialog.Show("Ошибка", "Выбранный элемент не является стеной.");
            this.IsEnabled = true;
        }

        // Метод для обработки выбора и размещения отверстий
        private void ProcessHolePlacement(UIDocument uiDoc, Element wall, bool isLinked)
        {
            List<string> settings = DockableManagerFormSettings.LoadSettings();

            if (settings == null)
            {
                var holeWindow = new sChoiseHole(_uiApp, wall, isLinked, userFullName, departmentName);
                holeWindow.ShowDialog();
            }
            else if (settings[2] != "Не выбрано" && settings[3] != "Не выбрано" && settings[4] != "Не выбрано")
            {
                _ExternalEventHandler.Instance.Raise((app) =>
                {
                    PlaceHoleOnWallCommand.Execute(app, userFullName, departmentName, wall, isLinked, departmentName, settings[3], settings[4]);
                });
            }
            else
            {
                var holeWindow = new sChoiseHole(_uiApp, wall, isLinked, userFullName, departmentName);
                holeWindow.ShowDialog();
            }
        }









        // XAML. Обработчик для кнопки "Настройка плагина"
        public void HolePluginSettings(object sender, RoutedEventArgs e)
        {
            var dockableManagerFormSettings = new DockableManagerFormSettings(userFullName, departmentName);
            dockableManagerFormSettings.ShowDialog();
        }

        // XAML. Вывод информации об отверстиях
        private void StatusButton_Click(object sender, RoutedEventArgs e)
        {
            Button clickedButton = sender as Button;
            if (clickedButton == null) return;

            // Очищаем панель
            InfoHolePanel.Children.Clear();
            InfoHolePanel.RowDefinitions.Clear();

            // Создаем вспомогательные элементы
            ScrollViewer scrollViewer = new ScrollViewer{VerticalScrollBarVisibility = ScrollBarVisibility.Auto};
            StackPanel holeListPanel = new StackPanel{};
            scrollViewer.Content = holeListPanel;
            InfoHolePanel.Children.Add(scrollViewer);

            // Получаем список всех сообщений для отверстий
            Document doc = _uiApp.ActiveUIDocument.Document;
            List<ElementId> familyInstanceIds = _iDataProcessor.GetFamilyInstanceIds(doc);
            List<List<string>> holeTaskMessages = _iDataProcessor.GetHoleLastTaskMessages(doc, familyInstanceIds);

            // Определяем нужный статус на основе нажатой кнопки
            string selectedStatus = null;
            if (clickedButton.Name == "ButtonNoneStatus") selectedStatus = "Без статуса";
            else if (clickedButton.Name == "ButtonOKStatus") selectedStatus = "Утверждено";
            else if (clickedButton.Name == "ButtonWarningStatus") selectedStatus = "Подтверждение";
            else if (clickedButton.Name == "ButtonErrorStatus") selectedStatus = "Ошибки";

            // Фильтруем сообщения по выбранному статусу
            List<List<string>> filteredMessages = null;

            if (departmentName == "BIM") 
            {
                filteredMessages = holeTaskMessages
                    .Where(messageParts => (messageParts[10] == selectedStatus))
                    .OrderByDescending(messageParts => messageParts[1] == userFullName)
                    .ThenBy(messageParts =>
                    {
                        string departmentIn = messageParts[3];
                        if (departmentIn == "АР") return 0;
                        if (departmentIn == "КР") return 1;
                        if (departmentIn == "ОВиК") return 2;
                        if (departmentIn == "ВК") return 3;
                        if (departmentIn == "ЭОМ") return 4;
                        if (departmentIn == "СС") return 5;
                        return 6;

                    })
                    .ThenBy(messageParts =>
                    {
                        string departmentFrom = messageParts[4];
                        if (departmentFrom == "АР") return 0;
                        if (departmentFrom == "КР") return 1;
                        if (departmentFrom == "ОВиК") return 2;
                        if (departmentFrom == "ВК") return 3;
                        if (departmentFrom == "ЭОМ") return 4;
                        if (departmentFrom == "СС") return 5;
                        return 6;
                    })
                    .ToList();
            }
            else
            {
                if (selectedStatus == "Без статуса") 
                {
                    filteredMessages = holeTaskMessages
                        .Where(messageParts => (messageParts[10] == selectedStatus) && (messageParts[3] == departmentName))
                        .OrderByDescending(messageParts => messageParts[1] == userFullName)
                        .ThenBy(messageParts =>
                        {
                            string departmentFrom = messageParts[4];
                            if (departmentFrom == "АР") return 0;
                            if (departmentFrom == "КР") return 1;
                            if (departmentFrom == "ОВиК") return 2;
                            if (departmentFrom == "ВК") return 3;
                            if (departmentFrom == "ЭОМ") return 4;
                            if (departmentFrom == "СС") return 5;
                            return 6;
                        })
                        .ToList();
                }
                else if (selectedStatus == "Подтверждение")
                {
                    filteredMessages = holeTaskMessages
                    .Where(messageParts => (messageParts[10] == selectedStatus) && ((messageParts[3] == departmentName) || (messageParts[4] == departmentName)))                    
                    .OrderByDescending(messageParts => messageParts[3] == departmentName)
                    .OrderByDescending(messageParts => messageParts[1] == userFullName)
                    .ThenBy(messageParts => messageParts[3] != departmentName) 
                    .ThenBy(messageParts =>
                    {
                        string departmentIn = messageParts[3];
                        if (departmentIn == "АР") return 0;
                        if (departmentIn == "КР") return 1;
                        if (departmentIn == "ОВиК") return 2;
                        if (departmentIn == "ВК") return 3;
                        if (departmentIn == "ЭОМ") return 4;
                        if (departmentIn == "СС") return 5;
                        return 6;
                    })
                    .ThenBy(messageParts =>
                    {
                        string departmentFrom = messageParts[4];
                        if (departmentFrom == "АР") return 0;
                        if (departmentFrom == "КР") return 1;
                        if (departmentFrom == "ОВиК") return 2;
                        if (departmentFrom == "ВК") return 3;
                        if (departmentFrom == "ЭОМ") return 4;
                        if (departmentFrom == "СС") return 5;
                        return 6;
                    })
                    .ToList();
                }
                else if (selectedStatus == "Ошибки")
                {
                    filteredMessages = holeTaskMessages
                    .Where(messageParts => (messageParts[10] == selectedStatus) && ((messageParts[3] == departmentName) || (messageParts[4] == departmentName)))
                    .OrderByDescending(messageParts => messageParts[4] == departmentName)
                    .OrderByDescending(messageParts => messageParts[1] == userFullName)
                    .ThenBy(messageParts => messageParts[4] != departmentName)
                    .ThenBy(messageParts =>
                    {
                        string departmentFrom = messageParts[4];
                        if (departmentFrom == "АР") return 0;
                        if (departmentFrom == "КР") return 1;
                        if (departmentFrom == "ОВиК") return 2;
                        if (departmentFrom == "ВК") return 3;
                        if (departmentFrom == "ЭОМ") return 4;
                        if (departmentFrom == "СС") return 5;
                        return 6;
                    })
                    .ThenBy(messageParts =>
                    {
                        string departmentIn = messageParts[3];
                        if (departmentIn == "АР") return 0;
                        if (departmentIn == "КР") return 1;
                        if (departmentIn == "ОВиК") return 2;
                        if (departmentIn == "ВК") return 3;
                        if (departmentIn == "ЭОМ") return 4;
                        if (departmentIn == "СС") return 5;
                        return 6;
                    })                   
                    .ToList();
                }
                else if (selectedStatus == "Утверждено")
                {
                    filteredMessages = holeTaskMessages
                    .Where(messageParts => (messageParts[10] == selectedStatus) && ((messageParts[3] == departmentName) || (messageParts[4] == departmentName)))
                    .OrderByDescending(messageParts => messageParts[4] == departmentName)
                    .OrderByDescending(messageParts => messageParts[1] == userFullName)
                    .ThenBy(messageParts => messageParts[4] != departmentName)
                    .ThenBy(messageParts =>
                    {
                        string departmentFrom = messageParts[4];
                        if (departmentFrom == "АР") return 0;
                        if (departmentFrom == "КР") return 1;
                        if (departmentFrom == "ОВиК") return 2;
                        if (departmentFrom == "ВК") return 3;
                        if (departmentFrom == "ЭОМ") return 4;
                        if (departmentFrom == "СС") return 5;
                        return 6;
                    })
                    .ThenBy(messageParts =>
                    {
                        string departmentIn = messageParts[3];
                        if (departmentIn == "АР") return 0;
                        if (departmentIn == "КР") return 1;
                        if (departmentIn == "ОВиК") return 2;
                        if (departmentIn == "ВК") return 3;
                        if (departmentIn == "ЭОМ") return 4;
                        if (departmentIn == "СС") return 5;
                        return 6;
                    })
                    
                    .ToList();
                }
            }

            foreach (var messageParts in filteredMessages)
            {
                // Разбиваем messageParts на отдельные переменные
                string data = messageParts[0];
                string name = messageParts[1];

                string departament = messageParts[2];
                string statusIO = messageParts[11];

                string departamentFrom = messageParts[3];
                char statusI = statusIO[0];
                string departamentIn = messageParts[4];
                char statusO = statusIO[1];

                string holeID = messageParts[5];
                string holeName = messageParts[6];
                string wallID = messageParts[7];
                string sEllementID = messageParts[8];

                // Содержимое TextBlock для newButton
                TextBlock textBlock = new TextBlock
                {
                    TextWrapping = TextWrapping.Wrap
                };

                textBlock.Inlines.Add(new Run($"{data}. ") { FontWeight = FontWeights.Bold, Foreground = Brushes.MediumBlue });

                var nameRun = new Run(name) { FontWeight = FontWeights.Bold };
                if (name == userFullName)
                {
                    nameRun.Foreground = Brushes.Purple;
                    nameRun.TextDecorations = TextDecorations.Underline;
                }
                else
                {
                    nameRun.Foreground = Brushes.Black;
                }
                textBlock.Inlines.Add(nameRun);

                textBlock.Inlines.Add(new Run(" ("));

                var departamentFromRun = new Run(departamentFrom) { FontWeight = FontWeights.Bold };
                if (departamentFrom == departmentName)
                {
                    departamentFromRun.Foreground = Brushes.MediumBlue;
                }
                else
                {
                    departamentFromRun.Foreground = Brushes.DimGray;
                }
                textBlock.Inlines.Add(departamentFromRun);

                textBlock.Inlines.Add(new Run(" -> "));

                var departamentInRun = new Run(departamentIn) { FontWeight = FontWeights.Bold };
                if (departamentIn == departmentName)
                {
                    departamentInRun.Foreground = Brushes.MediumBlue;
                }
                else
                {
                    departamentInRun.Foreground = Brushes.DimGray;
                }
                textBlock.Inlines.Add(departamentInRun);

                textBlock.Inlines.Add(new Run(").\n"));

                textBlock.Inlines.Add(new Run(holeName) { FontWeight = FontWeights.Bold, FontSize = 11, FontStyle = FontStyles.Italic });
                textBlock.Inlines.Add(new Run($" ({holeID}).\n") { FontSize = 11, FontStyle = FontStyles.Italic });
                textBlock.Inlines.Add(new Run("Стена: ") { FontWeight = FontWeights.Bold, FontSize = 11, FontStyle = FontStyles.Italic });
                textBlock.Inlines.Add(new Run($"{wallID}. ") { FontSize = 11, FontStyle = FontStyles.Italic });
                textBlock.Inlines.Add(new Run("Элементы в отверстии: ") { FontWeight = FontWeights.Bold, FontSize = 11, FontStyle = FontStyles.Italic });
                textBlock.Inlines.Add(new Run($"{sEllementID}.") { FontSize = 11, FontStyle = FontStyles.Italic });

                Button newButton = new Button
                {
                    Content = textBlock,
                    Height = 60,
                    BorderThickness = new Thickness(2),
                    HorizontalAlignment = HorizontalAlignment.Stretch,
                    HorizontalContentAlignment = HorizontalAlignment.Left,
                    Margin = new Thickness(2, 2, 2, 0),
                    Padding = new Thickness(10, 0, 0, 0)
                };

                // Определяем цвет newButton на основе статуса (нажатой кнопки)
                if (selectedStatus == "Без статуса") 
                {
                    newButton.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(235, 235, 235));
                }
                else if (selectedStatus == "Подтверждение")
                {
                    newButton.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(250, 245, 180));

                    if (departamentFrom == departmentName)
                    {
                        newButton.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(250, 245, 135));
                        newButton.BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(189, 183, 107));                         
                    }
                }
                else if (selectedStatus == "Ошибки") 
                {
                    newButton.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 190, 175));

                    if (departamentIn == departmentName)
                    {
                        newButton.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(253, 151, 119));
                        newButton.BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(178, 34, 34));
                    }
                }
                else if (selectedStatus == "Утверждено") 
                {
                    newButton.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 255, 210));
                    newButton.BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(46, 139, 87));
                }
                 
                // Добавляем обработчик нажатия
                newButton.Click += (s, ev) =>
                {
                    // Очищаем панель
                    InfoHolePanel.Children.Clear();
                    InfoHolePanel.RowDefinitions.Clear();

                    // Прокладываем Grid
                    InfoHolePanel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Общая информация
                    InfoHolePanel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Панель действия
                    InfoHolePanel.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Чат-область
                    InfoHolePanel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Поле добавления сообщения

                    ///////////////////// Блок 1. Базовая информация
                    TextBlock generalInfoTextBlock = new TextBlock
                    {
                        TextWrapping = TextWrapping.Wrap,
                        Padding = new Thickness(8)
                    };

                    // Заполение базовой информации
                    generalInfoTextBlock.Inlines.Add(new Run(holeName) { FontWeight = FontWeights.Bold });
                    generalInfoTextBlock.Inlines.Add(new Run($" ({holeID})\n"));

                    Brush statusColor = Brushes.Gray;
                    string statusText = messageParts[10];

                    switch (statusText)
                    {
                        case "Без статуса": statusColor = Brushes.Gray; break;
                        case "Утверждено": statusColor = Brushes.Green; break;
                        case "Подтверждение": statusColor = Brushes.DarkGoldenrod; break;
                        case "Ошибки": statusColor = Brushes.Red; break;
                    }

                    generalInfoTextBlock.Inlines.Add(new Run(statusText) { FontWeight = FontWeights.Bold, Foreground = statusColor });
                    generalInfoTextBlock.Inlines.Add(new Run($"\n"));
                    generalInfoTextBlock.Inlines.Add(new Run("Стена: ") { FontWeight = FontWeights.Bold });
                    generalInfoTextBlock.Inlines.Add(new Run($"{wallID}\n"));
                    generalInfoTextBlock.Inlines.Add(new Run("Элементы в отверстии: ") { FontWeight = FontWeights.Bold });
                    generalInfoTextBlock.Inlines.Add(new Run($"{sEllementID}"));

                    System.Windows.Controls.Grid.SetRow(generalInfoTextBlock, 0);
                    InfoHolePanel.Children.Add(generalInfoTextBlock);

                    ///////////////////// Блок 2. Изменение статуса отверстия
                    StackPanel decisionPanel = new StackPanel
                    {
                        Orientation = Orientation.Horizontal,
                        HorizontalAlignment = HorizontalAlignment.Left,
                        Margin = new Thickness(2, 0, 0, 0),
                    };

                    // Кнопки "Да"/"Нет"
                    Button yesButton = new Button
                    {
                        Content = "✔️",
                        Width = 30,
                        Height = 30,
                        Background = Brushes.Green,
                        Foreground = Brushes.White,
                        Margin = new Thickness(6, 0, 15, 0)
                    };
                    Button noButton = new Button
                    {
                        Content = "❌",
                        Width = 30,
                        Height = 30,
                        Background = Brushes.Red,
                        Foreground = Brushes.White,
                        Margin = new Thickness(0, 0, 15, 0)
                    };

                    yesButton.Click += (si, evi) =>
                        {
                            if (!int.TryParse(holeID, out int holeElementId))
                            {
                                return;
                            }

                            Element element = doc.GetElement(new ElementId(holeElementId));
                            FamilyInstance holeInstance = element as FamilyInstance;

                            string statusNext = "Без статуса";
                            string statusIONext = "22";

                            if (selectedStatus == "Без статуса")
                            {
                                statusNext = "Подтверждение";
                                statusIONext = "10";
                            }
                            else if (selectedStatus == "Подтверждение")
                            {
                                statusNext = "Утверждено";
                                statusIONext = "11";
                            }
                            else if (selectedStatus == "Ошибки")
                            {
                                statusNext = "Без статуса";
                                statusIONext = "32";
                            }

                            ExtensibleStorageHelper.AddChatMessage(
                                    holeInstance,
                                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                                    userFullName,
                                    departmentName,
                                    departamentFrom,
                                    departamentIn,
                                    wallID,
                                    sEllementID,
                                    statusNext,
                                    statusIONext,
                                    $"Смена статуса на отверстия `{statusNext}`"
                                    );

                            InfoHolePanel.Children.Clear();
                            InfoHolePanel.RowDefinitions.Clear();
                            UpdateStatusCounts();

                            TaskDialog.Show("Информация", $"{holeName} ({holeID}). Статус обновлён на ``{statusNext}``");
                        };

                    noButton.Click += (si, evi) =>
                    {
                        if (!int.TryParse(holeID, out int holeElementId))
                        {
                            return;
                        }

                        Element element = doc.GetElement(new ElementId(holeElementId));
                        FamilyInstance holeInstance = element as FamilyInstance;

                        string statusNext = "Без статуса";
                        string statusIONext = "22";

                        if (selectedStatus == "Подтверждение")
                        {
                            statusNext = "Ошибки";
                            statusIONext = "12";
                        }
                        else if (selectedStatus == "Утверждено")
                        {
                            statusNext = "Без статуса";
                            statusIONext = "31";
                        }

                        ExtensibleStorageHelper.AddChatMessage(
                                holeInstance,
                                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                                userFullName,
                                departmentName,
                                departamentFrom,
                                departamentIn,
                                wallID,
                                sEllementID,
                                statusNext,
                                statusIONext,
                                $"Смена статуса на отверстия ``{statusNext}``"
                                );

                        InfoHolePanel.Children.Clear();
                        InfoHolePanel.RowDefinitions.Clear();
                        UpdateStatusCounts();

                        TaskDialog.Show("Информация", $"Статус отверстия {holeID} обновлён на ``{statusNext}``");
                    };

                    // Функция для получения цвета по статусу
                    Brush GetStatusColor(char status)
                    {
                        if (status == '0') return Brushes.LightGray;
                        if (status == '1') return Brushes.LightGreen;
                        if (status == '2') return Brushes.LightPink;
                        if (status == '3') return Brushes.Violet;
                        return Brushes.LightGray;
                    }

                    TextBlock taskITextBlock = new TextBlock
                    {
                        Text = $"{departamentFrom}",
                        Width = 55,
                        Background = GetStatusColor(statusI),
                        Padding = new Thickness(7),
                        TextAlignment = TextAlignment.Center,
                        Margin = new Thickness(0, 0, 4, 0)
                    };

                    TextBlock taskOTextBlock = new TextBlock
                    {
                        Text = $"{departamentIn}",
                        Width = 55,
                        Background = GetStatusColor(statusO),
                        Padding = new Thickness(7),
                        TextAlignment = TextAlignment.Center
                    };

                    // Добавляем элементы в панель
                    if (departmentName != "BIM")
                    {
                        if (selectedStatus == "Без статуса")
                        {
                            decisionPanel.Children.Add(yesButton);
                        }
                        else if (selectedStatus == "Подтверждение")
                        {
                            if (departmentName != departamentFrom)
                            {
                                yesButton.Margin = new Thickness(6, 0, 3, 0);
                                decisionPanel.Children.Add(yesButton);
                                decisionPanel.Children.Add(noButton);
                            }
                            else
                            {
                                decisionPanel.Margin = new Thickness(8, 0, 0, 0);
                            }
                        }
                        else if (selectedStatus == "Ошибки")
                        {
                            if (departmentName == departamentFrom)
                            {
                                decisionPanel.Children.Add(yesButton);
                            }
                            else
                            {
                                decisionPanel.Margin = new Thickness(8, 0, 0, 0);
                            }
                        }
                        else if (selectedStatus == "Утверждено")
                        {
                            if (departmentName == departamentIn || departmentName == departamentFrom || userFullName == name)
                            {
                                noButton.Margin = new Thickness(6, 0, 15, 0);
                                decisionPanel.Children.Add(noButton);
                            }
                            else
                            {
                                decisionPanel.Margin = new Thickness(8, 0, 0, 0);
                            }
                        }
                    }
                    else
                    {
                        decisionPanel.Margin = new Thickness(8, 0, 0, 0);
                    }

                    decisionPanel.Children.Add(taskITextBlock);
                    decisionPanel.Children.Add(taskOTextBlock);

                    System.Windows.Controls.Grid.SetRow(decisionPanel, 1);
                    InfoHolePanel.Children.Add(decisionPanel);

                    ///////////////////// Блок 3. Чат и ScrollViewer
                    ScrollViewer messagesScrollViewer = new ScrollViewer
                    {
                        VerticalScrollBarVisibility = ScrollBarVisibility.Visible,
                        HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                        Margin = new Thickness(5, 8, 5, 8)
                    };

                    StackPanel messagesPanel = new StackPanel
                    {
                        VerticalAlignment = VerticalAlignment.Stretch
                    };

                    // Переменная для хранения координат предыдущего сообщения
                    string previousCoordinates = null;

                    foreach (List<string> fullHoleInfoParts in _iDataProcessor.GetHoleTaskMessages(doc, holeID))
                    {
                        if (fullHoleInfoParts.Count > 10)
                        {
                            string mDate = fullHoleInfoParts[0];
                            string mName = fullHoleInfoParts[1];
                            string mDepartmentFrom = fullHoleInfoParts[3];
                            string mDepartmentTo = fullHoleInfoParts[4];
                            string mCoordinates = fullHoleInfoParts[9];
                            string mMessageText = fullHoleInfoParts[12];

                            TextBlock messageTextBlock = new TextBlock { TextWrapping = TextWrapping.Wrap };

                            messageTextBlock.Inlines.Add(new Run($"{mDate}") { FontWeight = FontWeights.Bold, Foreground = Brushes.BlueViolet });
                            messageTextBlock.Inlines.Add(new Run($" | {mName} ({mDepartmentFrom} → {mDepartmentTo})\n"));

                            messageTextBlock.Inlines.Add(new Run("Координаты: ") { FontWeight = FontWeights.Bold });

                            // Проверяем, изменились ли координаты с прошлого сообщения
                            Brush coordinatesColor = previousCoordinates != null && previousCoordinates != mCoordinates
                                ? Brushes.Red
                                : Brushes.Black;

                            messageTextBlock.Inlines.Add(new Run($"{mCoordinates}\n") { Foreground = coordinatesColor });

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

                            // Обновляем previousCoordinates для следующей итерации
                            previousCoordinates = mCoordinates;

                            messagesPanel.Children.Add(messageBorder);
                        }
                    }

                    messagesScrollViewer.Content = messagesPanel;
                    System.Windows.Controls.Grid.SetRow(messagesScrollViewer, 2);
                    InfoHolePanel.Children.Add(messagesScrollViewer);

                    ///////////////////// Блок 4. Блок добавления комментария
                    System.Windows.Controls.Grid sendMessagesPanel = new System.Windows.Controls.Grid
                    {
                        Margin = new Thickness(5, 5, 5, 5),
                        VerticalAlignment = VerticalAlignment.Bottom
                    };

                    sendMessagesPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                    sendMessagesPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(30) });

                    System.Windows.Controls.TextBox commentTextBox = new System.Windows.Controls.TextBox
                    {
                        Height = 45,
                        VerticalAlignment = VerticalAlignment.Center,
                        Padding = new Thickness(5, 5, 5, 5),
                        Margin = new Thickness(5, 0, 3, 0),
                        AcceptsReturn = true,
                        TextWrapping = TextWrapping.Wrap,
                        VerticalScrollBarVisibility = ScrollBarVisibility.Auto
                    };

                    System.Windows.Controls.Grid.SetColumn(commentTextBox, 0);

                    Button sendButton = new Button
                    {
                        Content = "🔼",
                        Height = 45,
                        Width = 30,
                        VerticalAlignment = VerticalAlignment.Center,
                        HorizontalAlignment = HorizontalAlignment.Stretch,
                        IsEnabled = false
                    };

                    commentTextBox.TextChanged += (si, evi) =>
                    {
                        sendButton.IsEnabled = !string.IsNullOrWhiteSpace(commentTextBox.Text);
                    };

                    sendButton.Click += (si, evi) =>
                    {
                        if (!int.TryParse(holeID, out int holeElementId))
                        {
                            return;
                        }

                        Element element = doc.GetElement(new ElementId(holeElementId));
                        FamilyInstance holeInstance = element as FamilyInstance;

                        string commentText = commentTextBox.Text;

                        ExtensibleStorageHelper.AddChatMessage(
                            holeInstance,
                            DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                            userFullName,
                            departmentName,
                            departamentFrom,
                            departamentIn,
                            wallID,
                            sEllementID,
                            statusText,
                            statusIO,
                            commentText
                            );

                        newButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                    };

                    System.Windows.Controls.Grid.SetColumn(sendButton, 1);

                    sendMessagesPanel.Children.Add(commentTextBox);
                    sendMessagesPanel.Children.Add(sendButton);

                    System.Windows.Controls.Grid.SetRow(sendMessagesPanel, 3);
                    InfoHolePanel.Children.Add(sendMessagesPanel);

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

                holeListPanel.Children.Add(newButton);
            }
        }
    }

    // Передача данных статусов в названия кнопок
    public class ButtonDataViewModel : INotifyPropertyChanged
    {
        // Первичная прогрузка
        private string _noneStatusButtonText = "❓  Без статуса: ОШИБКА!";
        private string _approvedButtonText = "✔️  Утверждено: ОШИБКА!";
        private string _warningButtonText = "⚠️  Подтверждение: ОШИБКА!";
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
            WarningButtonText = $"⚠️  Подтверждение: {statusCounts[2]}";
            ErrorButtonText = $"❌  Ошибки: {statusCounts[3]}";
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    // Фильтр для выбора стен и линков
    public class WallAndLinkSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is Wall || elem is RevitLinkInstance;
        }

        public bool AllowReference(Reference reference, XYZ position) => false;
    }
}