using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Markup;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using System.Xml.Linq;
using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_CoordiantorAI.Common;
using KPLN_CoordiantorAI.ExternalModel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Windows.Controls.Primitives;
using static KPLN_CoordiantorAI.ExternalModel.Commands;
using WpfGrid = System.Windows.Controls.Grid;
using WpfTextBox = System.Windows.Controls.TextBox;
using Control = System.Windows.Controls.Control;

namespace KPLN_CoordiantorAI.Forms
{
    public partial class ExternalModelWindow : Window
    {
        public ExternalModelWindow(
            Autodesk.Revit.DB.Document document,
            UIDocument uiDocument,
            ConnectionType connectionType,
            ExternalModelSettings settings)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));
            if (uiDocument == null)
                throw new ArgumentNullException(nameof(uiDocument));

            InitializeComponent();

            ExternalModelControl control = new ExternalModelControl(document, uiDocument, connectionType, settings);
            ModelContent.Content = control;
            Title = control.DisplayTitle;
        }
    }

    /// <summary>
    /// Рабочая область взаимодействия с внешней моделью.
    /// </summary>
    public class ExternalModelControl : UserControl
    {
        private Autodesk.Revit.DB.Document _doc;
        private UIDocument _uiDoc;
        private ChatLogger _logger;
        private ConnectionType _connectionType;
        private ExternalModelSettings _settings;

        private int _lastCacheHit = 0;
        private int _lastCacheMiss = 0;
        private int _lastCompletion = 0;
        private int _lastTotal = 0;


        private readonly HttpClient _httpClient = new HttpClient();
        public List<object> ChatHistoryMessages { get; } = new List<object>();

        private Border _typingIndicator;
        private DispatcherTimer _typingTimer;
        private TextBlock TitleTextBlock;
        private ScrollViewer ChatScrollViewer;
        private StackPanel ChatHistory;
        private WpfTextBox InputTextBox;
        private Button SendButton;

        public string DisplayTitle { get; private set; }

        public ExternalModelControl(
            Autodesk.Revit.DB.Document document,
            UIDocument uiDocument,
            ConnectionType connectionType,
            ExternalModelSettings settings)
        {
            InitializeModelLayout();
            _doc = document;
            _uiDoc = uiDocument;
            _connectionType = connectionType;
            _settings = settings ?? new ExternalModelSettings();

            // Инициализация логгера
            _logger = new ChatLogger(_settings.LogFolder);

            // Анимация точек
            _typingTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(400) };
            _typingTimer.Tick += TypingTimer_Tick;

            SetupEventHandlers();

            // Показываем пользователю, какой режим активен
            DisplayTitle = GetModelTitle();
            TitleTextBlock.Text = DisplayTitle;
        }

        // Показываем пользователю, какой режим активен
        private string GetModelTitle()
        {
            string mode = _connectionType == ConnectionType.OnlineAPI
                ? "Online (DeepSeek API)"
                : "Local (LM Studio)";

            return $"Работа с моделью - {_doc.Title} - {mode}";
        }

        private void InitializeModelLayout()
        {
            MinHeight = 480;
            MinWidth = 680;
            Background = CreateBrush(32, 36, 45);
            FontFamily = new FontFamily("Segoe UI");

            WpfGrid root = new WpfGrid
            {
                Background = CreateBrush(32, 36, 45)
            };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            Border header = new Border
            {
                Background = CreateBrush(38, 43, 53),
                BorderBrush = CreateBrush(52, 59, 73),
                BorderThickness = new Thickness(0, 0, 0, 1),
                Padding = new Thickness(12, 10, 12, 10)
            };
            WpfGrid.SetRow(header, 0);

            TitleTextBlock = new TextBlock
            {
                Text = "Работа с моделью",
                FontSize = 16,
                FontWeight = FontWeights.SemiBold,
                Foreground = Brushes.White
            };
            header.Child = TitleTextBlock;
            root.Children.Add(header);

            ChatHistory = new StackPanel();
            ChatScrollViewer = new ScrollViewer
            {
                Content = ChatHistory,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Margin = new Thickness(10)
            };
            WpfGrid.SetRow(ChatScrollViewer, 1);
            root.Children.Add(ChatScrollViewer);

            WpfGrid inputGrid = new WpfGrid
            {
                Margin = new Thickness(10, 0, 10, 10)
            };
            inputGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            inputGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            WpfGrid.SetRow(inputGrid, 2);



            InputTextBox = new WpfTextBox
            {
                Language = XmlLanguage.GetLanguage("ru-RU"),

                Margin = new Thickness(0, 0, 10, 0),
                Padding = new Thickness(10, 8, 10, 8),

                // Фиксированная высота. Поле больше не растёт вверх.
                Height = 80,

                // Текст внутри начинается сверху
                VerticalContentAlignment = VerticalAlignment.Top,

                // Многострочный ввод
                AcceptsReturn = true,

                // Перенос строк
                TextWrapping = TextWrapping.Wrap,

                // Горизонтальный скролл не нужен, вертикальный появляется при переполнении
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,

                FontSize = 14,
                Foreground = Brushes.White,
                Background = CreateBrush(42, 48, 59),
                BorderBrush = CreateBrush(75, 85, 104),
                CaretBrush = Brushes.White
            };
            SpellCheck.SetIsEnabled(InputTextBox, true);


            WpfGrid.SetColumn(InputTextBox, 0);
            inputGrid.Children.Add(InputTextBox);

            SendButton = new Button
            {
                Content = "Отправить",
                Style = CreateSendButtonStyle(),
                VerticalAlignment = VerticalAlignment.Bottom
            };
            WpfGrid.SetColumn(SendButton, 1);
            inputGrid.Children.Add(SendButton);

            root.Children.Add(inputGrid);
            Content = root;
        }

        private static Brush CreateBrush(byte red, byte green, byte blue)
        {
            return new SolidColorBrush(System.Windows.Media.Color.FromRgb(red, green, blue));
        }

        private static Style CreateSendButtonStyle()
        {
            Style style = new Style(typeof(Button));

            style.Setters.Add(new Setter(Control.MinHeightProperty, 30.0));
            style.Setters.Add(new Setter(Control.MinWidthProperty, 92.0));
            style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(15, 10, 15, 10)));
            style.Setters.Add(new Setter(Control.BackgroundProperty, CreateBrush(44, 107, 237)));
            style.Setters.Add(new Setter(Control.ForegroundProperty, Brushes.White));
            style.Setters.Add(new Setter(Control.BorderBrushProperty, CreateBrush(44, 107, 237)));
            style.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(1)));

            FrameworkElementFactory borderFactory = new FrameworkElementFactory(typeof(Border));
            borderFactory.Name = "ButtonBorder";
            borderFactory.SetBinding(Border.BackgroundProperty, new System.Windows.Data.Binding("Background")
            {
                RelativeSource = RelativeSource.TemplatedParent
            });
            borderFactory.SetBinding(Border.BorderBrushProperty, new System.Windows.Data.Binding("BorderBrush")
            {
                RelativeSource = RelativeSource.TemplatedParent
            });
            borderFactory.SetBinding(Border.BorderThicknessProperty, new System.Windows.Data.Binding("BorderThickness")
            {
                RelativeSource = RelativeSource.TemplatedParent
            });
            borderFactory.SetBinding(Border.PaddingProperty, new System.Windows.Data.Binding("Padding")
            {
                RelativeSource = RelativeSource.TemplatedParent
            });

            FrameworkElementFactory contentFactory = new FrameworkElementFactory(typeof(ContentPresenter));
            contentFactory.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            contentFactory.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            contentFactory.SetValue(ContentPresenter.RecognizesAccessKeyProperty, true);

            borderFactory.AppendChild(contentFactory);

            ControlTemplate template = new ControlTemplate(typeof(Button));
            template.VisualTree = borderFactory;

            Trigger hoverTrigger = new Trigger
            {
                Property = UIElement.IsMouseOverProperty,
                Value = true
            };
            hoverTrigger.Setters.Add(new Setter(UIElement.OpacityProperty, 0.96, "ButtonBorder"));
            template.Triggers.Add(hoverTrigger);

            Trigger pressedTrigger = new Trigger
            {
                Property = ButtonBase.IsPressedProperty,
                Value = true
            };
            pressedTrigger.Setters.Add(new Setter(UIElement.OpacityProperty, 0.86, "ButtonBorder"));
            template.Triggers.Add(pressedTrigger);

            Trigger disabledTrigger = new Trigger
            {
                Property = UIElement.IsEnabledProperty,
                Value = false
            };
            disabledTrigger.Setters.Add(new Setter(UIElement.OpacityProperty, 0.55, "ButtonBorder"));
            template.Triggers.Add(disabledTrigger);

            style.Setters.Add(new Setter(Control.TemplateProperty, template));

            return style;
        }

        //Задать текст под загрузку ИИ
        private void TypingTimer_Tick(object sender, EventArgs e)
        {
            if (_typingIndicator?.Child is RichTextBox richTextBox)
            {
                var frames = new[]
                {
                    "ИИ печатает",
                    "ИИ печатает.",
                    "ИИ печатает..",
                    "ИИ печатает..."
                };

                _currentTypingFrame = (_currentTypingFrame + 1) % frames.Length;

                var paragraph = new Paragraph();
                paragraph.Inlines.Add(new Run(frames[_currentTypingFrame]));
                richTextBox.Document = new FlowDocument(paragraph);
            }

        }

        //показать загрузку ИИ
        private void ShowTypingIndicator()
        {
            if (_typingIndicator != null)
                ChatHistory.Children.Remove(_typingIndicator);

            _typingIndicator = CreateTypingIndicator();
            ChatHistory.Children.Add(_typingIndicator);
            ChatScrollViewer.ScrollToEnd();

            _currentTypingFrame = 0;
            _typingTimer.Start();  // ← Запуск анимации
        }

        //Скрыть загрузку ИИ
        private void HideTypingIndicator()
        {
            _typingTimer.Stop();
            if (_typingIndicator != null)
            {
                ChatHistory.Children.Remove(_typingIndicator);
                _typingIndicator = null;
            }
        }


        private int _currentTypingFrame = 0;

        //подписка на события при отправке сообщения (кнпока "Отправить"/Enter)
        private void SetupEventHandlers()
        {
            SendButton.Click += SendButton_Click;
            InputTextBox.KeyDown += InputTextBox_KeyDown;
            InputTextBox.PreviewKeyDown += InputTextBox_KeyDown;
        }

        //кнопка отправить сообщение
        private async void SendButton_Click(object sender, RoutedEventArgs e)
        {
            await SendMessage();
        }

        //позволяет оотправить сообщение через enter
        private void InputTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if ((e.Key == Key.Enter || e.Key == Key.Return) && Keyboard.Modifiers != ModifierKeys.Shift)
            {
                e.Handled = true;
                SendButton_Click(null, null);
            }
        }

        private bool IsSuccessResponse(string responseJson)
        {
            try
            {
                var obj = JObject.Parse(responseJson);
                return obj["error"] == null;
            }
            catch
            {
                return false;
            }
        }

        //процесс отправки сообщения
        public async Task SendMessage(string revitContext = "")
        {
            string userMessage = InputTextBox.Text.Trim();
            if (string.IsNullOrEmpty(userMessage)) return;

            DateTime requestTime = DateTime.Now;
            string requestModelName = GetCurrentRevitModelName();
            string requestViewName = GetCurrentRevitViewName();

            // Пользовательское сообщение
            var userMsg = new { role = "user", content = userMessage };
            ChatHistoryMessages.Add(userMsg);

            ChatHistory.Children.Add(CreateMessageBlock($"Вы: {userMessage}", true));
            InputTextBox.Clear();
            SendButton.IsEnabled = false;

            try
            {
                ShowTypingIndicator(); //начало анимации загрузки

                string response = await SendToOpenRouter(ChatHistoryMessages);

                if (!response.TrimStart().StartsWith("{"))
                {
                    ChatHistory.Children.Add(CreateMessageBlock($"❌ AI: {response}", false));


                    if (ChatHistoryMessages.Count > 0 && ChatHistoryMessages.Last().GetType().GetProperty("role")?.GetValue(ChatHistoryMessages.Last())?.ToString() == "tool")
                    {
                        ChatHistoryMessages.RemoveAt(ChatHistoryMessages.Count - 1);
                    }


                    return;
                }

                // Парсим ответ ИИ
                var responseJObject = JObject.Parse(response);

                while (true)
                {
                    var message = responseJObject["choices"]?[0]?["message"].ToObject<JObject>();
                    var toolCalls = message?["tool_calls"] as JArray;


                    ChatHistoryMessages.Add(message.ToObject<object>());


                    // Если есть tool_calls — выполняем
                    if (toolCalls != null && toolCalls.Count > 0)
                    {
                        foreach (JObject tc in toolCalls)
                        {
                            ProcessSingleToolCall(tc);
                        }

                        // После всех tools → следующий запрос к ИИ
                        response = await SendToOpenRouter(ChatHistoryMessages);
                        responseJObject = JObject.Parse(response);

                        if (!IsSuccessResponse(response))
                        {
                            // Если ошибка - показываем и выходим
                            ChatHistory.Children.Add(CreateMessageBlock($"❌ Ошибка после tools: {response}", false));
                            break;
                        }

                    }
                    else
                    {
                        // Нет больше tools — показываем финальный ответ
                        var finalContent = message?["content"]?.ToString() ?? "";

                        if (finalContent.TrimStart().StartsWith("{") && finalContent.Contains("\"name\":"))
                        {
                            ChatHistory.Children.Add(CreateMessageBlock("🤖 ИИ думает... (обработка команд)", false));
                            continue;  // Пропускаем RAW JSON
                        }

                        if (!string.IsNullOrEmpty(finalContent) && finalContent.Trim().Length > 10)
                        {
                            DateTime responseTime = DateTime.Now;
                            ChatHistory.Children.Add(CreateMessageBlock($"AI: {finalContent}", false));

                            // Ищем и удаляем всю цепочку вызовов инструментов
                            for (int i = 0; i < ChatHistoryMessages.Count; i++)
                            {
                                var msg = ChatHistoryMessages[i];
                                var msgType = msg.GetType();
                                string role = msgType.GetProperty("role")?.GetValue(msg)?.ToString();

                                if (role == "assistant")
                                {
                                    var toolCallsDel = msgType.GetProperty("tool_calls")?.GetValue(msg);
                                    if (toolCallsDel != null)
                                    {
                                        // Нашли сообщение с tool_calls — удаляем всё от него до конца
                                        int countToRemove = ChatHistoryMessages.Count - i;
                                        ChatHistoryMessages.RemoveRange(i, countToRemove);
                                        break;
                                    }
                                }
                            }

                            /// Добавляем чистый ответ ассистента (без tool_calls)
                            ChatHistoryMessages.Add(new { role = "assistant", content = finalContent });

                            _logger.LogWithTokens(
                                userMessage,
                                finalContent,
                                _lastCacheHit,
                                _lastCacheMiss,
                                _lastCompletion,
                                _lastTotal,
                                requestTime,
                                responseTime,
                                requestModelName,
                                requestViewName);
                            break;
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                ChatHistory.Children.Add(CreateMessageBlock($"Ошибка: {ex.Message}", false));
                // Логируем ошибку (вопрос пользователя и текст ошибки)
                _logger.Log(userMessage, $"ОШИБКА: {ex.Message}", requestTime, DateTime.Now, requestModelName, requestViewName);
            }
            finally
            {
                HideTypingIndicator();  //конец анимации загрузки
                SendButton.IsEnabled = true;
                ChatScrollViewer.ScrollToEnd();
            }
        }



        private string GetCurrentRevitModelName()
        {
            if (_doc == null)
                return "Документ Revit не найден";

            if (!string.IsNullOrWhiteSpace(_doc.Title))
                return _doc.Title;

            return string.IsNullOrWhiteSpace(_doc.PathName) ? "Без имени" : System.IO.Path.GetFileNameWithoutExtension(_doc.PathName);
        }

        private string GetCurrentRevitViewName()
        {
            Autodesk.Revit.DB.View activeView = _doc == null ? null : _doc.ActiveView;
            return activeView == null || string.IsNullOrWhiteSpace(activeView.Name)
                ? "Активный вид не найден"
                : activeView.Name;
        }



        private void ProcessSingleToolCall(JObject toolCall)
        {
            string toolName = toolCall["function"]?["name"]?.ToString();
            string toolCallId = toolCall["id"]?.ToString() ?? Guid.NewGuid().ToString();

            string toolResult = "";
            var argsJson = toolCall["function"]?["arguments"]?.ToString() ?? "{}";
            var argsObj = JObject.Parse(argsJson);  // ← ТУТ объявляем!

            switch (toolName)
            {
                case "get_active_view_in_revit":
                    var viewInfo = Commands.GetActiveViewInfo(_doc);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(viewInfo);
                    //ChatHistory.Children.Add(CreateMessageBlock($"✅ Активный вид: {viewInfo.ViewName} (ID: {viewInfo.ViewId}, Тип: {viewInfo.ViewType})", false));
                    break;

                case "get_all_elements_shown_in_view":
                    var viewIdParam = toolCall["function"]?["arguments"]?.ToString();
                    var args = JObject.Parse(viewIdParam ?? "{}");
                    int viewId = args["viewOrSheetId"]?.Value<int>() ?? IDHelper.ElIdInt(_doc.ActiveView.Id);
                    var elementsResult = Commands.GetAllElementsShownInView(_doc, viewId);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(elementsResult);

                    break;

                case "get_category_by_keyword":
                    var keywordParam = toolCall["function"]?["arguments"]?.ToString();
                    var keywordArgs = JObject.Parse(keywordParam ?? "{}");
                    string keyword = keywordArgs["keyword"]?.Value<string>() ?? "";
                    var categories = Commands.GetCategoryByKeyword(_doc, keyword);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(categories);

                    //foreach (var cat in categories.Take(3))
                    //    ChatHistory.Children.Add(CreateMessageBlock($"   • {cat.Name} (ID: {cat.Id})", false));
                    break;

                case "get_elements_by_category":
                    var catParam = toolCall["function"]?["arguments"]?.ToString();
                    var catArgs = JObject.Parse(catParam ?? "{}");
                    int categoryId = catArgs["categoryId"]?.Value<int>() ?? 0;
                    var elementIds = Commands.GetElementsByCategory(_doc, categoryId);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(elementIds);

                    break;

                case "get_model_categories":
                    var allCats = Commands.GetModelCategories(_doc);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(allCats);
                    break;

                case "get_categories_from_elementids":
                    argsJson = toolCall["function"]?["arguments"]?.ToString();
                    argsObj = JObject.Parse(argsJson ?? "{}");
                    var idsToken = argsObj["list_elementIds"];
                    var ids = new List<int>();
                    if (idsToken is JArray arr)
                        ids = arr.Select(t => t.Value<int>()).ToList();
                    var catMap = Commands.GetCategoriesFromElementIds(_doc, ids);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(catMap);
                    break;

                case "get_element_types_for_elementids":
                    var typeArgsJson = toolCall["function"]?["arguments"]?.ToString();
                    var typeArgsObj = JObject.Parse(typeArgsJson ?? "{}");
                    var typeIdsToken = typeArgsObj["list_elementIds"];
                    ids = new List<int>();
                    if (typeIdsToken is JArray typeArr)
                        ids = typeArr.Select(t => t.Value<int>()).ToList();
                    var result = Commands.GetElementTypesForElementIds(_doc, ids);
                    var typeMap = result.GetType().GetProperty("type_ids")?.GetValue(result) as Dictionary<int, ElementTypeInfo>;
                    int count = (int)(result.GetType().GetProperty("count")?.GetValue(result) ?? 0);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result);
                    break;

                case "get_all_elementids_for_specific_type_ids":

                    var typeIdsToken_8 = argsObj["list_typeIds"];
                    var typeIds = new List<int>();
                    if (typeIdsToken_8 is JArray typeArray)
                        typeIds = typeArray.Select(t => t.Value<int>()).ToList();
                    var result_8 = Commands.GetAllElementIdsForSpecificTypeIds(_doc, typeIds);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_8);
                    break;

                case "get_all_used_families_in_model":

                    var familyResult = Commands.GetAllUsedFamiliesInModel(_doc);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(familyResult);
                    break;

                case "get_all_used_families_of_category":

                    int categoryId_9 = argsObj["categoryId"]?.Value<int>() ?? 0;
                    var familyResult_9 = Commands.GetAllUsedFamiliesOfCategory(_doc, categoryId_9);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(familyResult_9);
                    break;

                case "get_all_used_types_of_a_family":
                    string familyName = argsObj["familyName"]?.Value<string>() ?? "";
                    var result_11 = Commands.GetAllUsedTypesOfAFamily(_doc, familyName);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_11);
                    break;

                case "get_all_elements_of_specific_families":
                    var names = new List<string>();
                    if (argsObj["familyNames"] is JArray famArr)
                        names = famArr.Select(t => t.Value<string>()).ToList();
                    var result_12 = Commands.GetAllElementsOfSpecificFamilies(_doc, names);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_12);
                    break;

                case "get_parameters_from_elementid":
                    int elementId = argsObj["elementId"]?.Value<int>() ?? 0;
                    bool getIdValuesAsNames = argsObj["getIdValuesAsNames"]?.Value<bool>() ?? false;
                    var result_13 = Commands.GetParametersFromElementId(_doc, elementId, getIdValuesAsNames);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_13);
                    break;

                case "get_parameter_value_for_element_ids":
                    var ids_14 = new List<int>();
                    if (argsObj["list_elementIds"] is JArray arr_14)
                        ids = arr_14.Select(t => t.Value<int>()).ToList();
                    int idParameter = argsObj["idParameter"]?.Value<int>() ?? 0;
                    bool getIdValuesAsNames_14 = argsObj["getIdValuesAsNames"]?.Value<bool>() ?? false;
                    var result_14 = Commands.GetParameterValueForElementIds(_doc, ids_14, idParameter, getIdValuesAsNames_14);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_14);
                    break;

                case "get_all_additional_properties_from_elementid":
                    int elementId_15 = argsObj["elementId"]?.Value<int>() ?? 0;
                    var result_15 = Commands.GetAllAdditionalPropertiesFromElementId(_doc, elementId_15);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_15);
                    break;

                case "get_additional_property_for_all_elementids":
                    var ids_16 = new List<int>();
                    if (argsObj["list_elementIds"] is JArray arr_16)
                        ids = arr_16.Select(t => t.Value<int>()).ToList();
                    string propertyName = argsObj["propertyName"]?.Value<string>() ?? "";
                    var result_16 = Commands.GetAdditionalPropertyForAllElementIds(_doc, ids_16, propertyName);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_16);
                    break;

                case "get_location_for_element_ids":
                    var ids_17 = new List<int>();
                    if (argsObj["list_elementIds"] is JArray arr_17)
                        ids_17 = arr_17.Select(t => t.Value<int>()).ToList();

                    var result_17 = Commands.GetLocationForElementIds(_doc, ids_17);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_17);
                    break;

                case "get_boundingboxes_for_element_ids":
                    var ids_18 = new List<int>();
                    int? idSheet = null;
                    if (argsObj["list_elementIds"] is JArray arr_18)
                        ids = arr_18.Select(t => t.Value<int>()).ToList();

                    if (argsObj["idSheet"] != null)
                        idSheet = argsObj["idSheet"].Value<int?>();

                    var result_18 = Commands.GetBoundingBoxesForElementIds(_doc, ids_18, idSheet);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_18);
                    break;

                case "get_boundary_lines":
                    var ids_19 = new List<int>();
                    if (argsObj["list_elementIds"] is JArray arr_19)
                        ids_19 = arr_19.Select(t => t.Value<int>()).ToList();
                    var result_19 = Commands.GetBoundaryLines(_doc, ids_19);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_19);
                    break;

                case "get_room_boundary_lines":
                    var ids_room = new List<int>();
                    // Проверяем наличие параметра list_roomIds
                    if (argsObj["list_roomIds"] is JArray arr_room)
                        ids_room = arr_room.Select(t => t.Value<int>()).ToList();
                    // Вызываем метод получения границ помещений
                    var result_room = Commands.GetRoomBoundaryLines(_doc, ids_room);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_room);
                    break;

                case "get_host_id_for_element_ids":
                    var ids_20 = new List<int>();
                    if (argsObj["list_elementIds"] is JArray arr_20)
                        ids = arr_20.Select(t => t.Value<int>()).ToList();
                    var result_20 = Commands.GetHostIdForElementIds(_doc, ids_20);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_20);
                    break;

                case "get_object_classes_from_elementids":
                    var ids_21 = new List<int>();
                    if (argsObj["list_elementIds"] is JArray arr_21)
                        ids_21 = arr_21.Select(t => t.Value<int>()).ToList();
                    var result_21 = Commands.GetObjectClassesFromElementIds(_doc, ids_21);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_21);
                    break;

                case "get_material_layers_from_types":
                    var ids_22 = new List<int>();

                    if (argsObj["list_elementIds"] is JArray arr_22)
                        ids_22 = arr_22.Select(t => t.Value<int>()).ToList();

                    var result_22 = Commands.GetMaterialLayersFromTypes(_doc, ids_22);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_22);
                    break;

                case "get_model_file_info":
                    var result_fileInfo = Commands.GetModelFileInfo(_doc);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_fileInfo);
                    break;
#if R2020
                case "get_all_project_units":
                    var result_units = Commands.GetAllProjectUnits(_doc);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_units);
                    break;
#endif
                case "get_all_warnings_in_the_model":
                    var result_warnings = Commands.GetAllWarningsInTheModel(_doc);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_warnings);
                    break;

                case "get_all_workset_information":
                    var result_worksets = Commands.GetAllWorksetInformation(_doc);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_worksets);
                    break;

                case "get_worksets_from_elementids":
                    var ids_workset = new List<int>();
                    if (argsObj["list_elementIds"] is JArray arr_workset)
                        ids_workset = arr_workset.Select(t => t.Value<int>()).ToList();

                    var result_workset = Commands.GetWorksetsFromElementIds(_doc, ids_workset);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_workset);
                    break;

                case "get_worksharing_information_for_element_ids":
                    var ids_worksharing = new List<int>();
                    if (argsObj["list_elementIds"] is JArray arr_worksharing)
                        ids_worksharing = arr_worksharing.Select(t => t.Value<int>()).ToList();

                    var result_worksharing = Commands.GetWorksharingInformationForElementIds(_doc, ids_worksharing);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_worksharing);
                    break;

                case "get_user_selection_in_revit":
                    var result_selection = Commands.GetUserSelectionInRevit(_doc, _uiDoc);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_selection);
                    break;

                case "set_user_selection_in_revit":
                    var ids_selection = new List<int>();
                    if (argsObj["list_elementIds"] is JArray arr_selection)
                        ids_selection = arr_selection.Select(t => t.Value<int>()).ToList();

                    var result_selection_30 = Commands.SetUserSelectionInRevit(_doc, _uiDoc, ids_selection);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_selection_30);
                    break;

                case "get_graphic_overrides_for_element_ids_in_view":
                    var ids_overrides = new List<int>();
                    if (argsObj["list_elementIds"] is JArray arr_overrides)
                        ids_overrides = arr_overrides.Select(t => t.Value<int>()).ToList();

                    int viewId_31 = argsObj["viewId"]?.Value<int>() ?? -1;

                    if (viewId_31 == -1)
                    {
                        toolResult = JsonConvert.SerializeObject(new { error = "Не указан viewId" });
                        break;
                    }

                    var result_overrides = Commands.GetGraphicOverridesForElementIdsInView(_doc, ids_overrides, viewId_31);
                    toolResult = JsonConvert.SerializeObject(result_overrides);
                    break;

                case "get_graphic_filters_applied_to_views":
                    var ids_filters = new List<int>();
                    if (argsObj["list_elementIds"] is JArray arr_filters)
                        ids_filters = arr_filters.Select(t => t.Value<int>()).ToList();

                    var result_filters = Commands.GetGraphicFiltersAppliedToViews(_doc, ids_filters);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_filters);
                    break;

                case "get_all_parameter_filters_in_model":
                    var result_allFilters = Commands.GetAllParameterFiltersInModel(_doc);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_allFilters);
                    break;

                case "get_graphic_overrides_view_filters":
                    var ids_filterOverrides = new List<int>();
                    if (argsObj["list_filterIds"] is JArray arr_filterOverrides)
                        ids_filterOverrides = arr_filterOverrides.Select(t => t.Value<int>()).ToList();
                    int viewIdFilter = argsObj["viewId"]?.Value<int>() ?? -1;
                    if (viewIdFilter == -1)
                    {
                        toolResult = JsonConvert.SerializeObject(new { error = "Не указан viewId" });
                        break;
                    }
                    var result_filterOverrides = Commands.GetGraphicOverridesViewFilters(_doc, ids_filterOverrides, viewIdFilter);
                    toolResult = JsonConvert.SerializeObject(result_filterOverrides);
                    break;

                case "get_category_visibility_overrides_in_view":
                    int viewIdForCategory = argsObj["viewId"]?.Value<int>() ?? -1;

                    if (viewIdForCategory == -1)
                    {
                        toolResult = JsonConvert.SerializeObject(new { error = "Не указан viewId" });
                        break;
                    }

                    var result_categoryOverrides = Commands.GetCategoryVisibilityOverridesInView(_doc, viewIdForCategory);
                    toolResult = JsonConvert.SerializeObject(result_categoryOverrides);
                    break;

                case "get_workset_visibility_in_view":
                    int viewIdForWorkset = argsObj["viewId"]?.Value<int>() ?? -1;

                    if (viewIdForWorkset == -1)
                    {
                        toolResult = JsonConvert.SerializeObject(new { error = "Не указан viewId" });
                        break;
                    }

                    var result_worksetVisibility = Commands.GetWorksetVisibilityInView(_doc, viewIdForWorkset);
                    toolResult = JsonConvert.SerializeObject(result_worksetVisibility);
                    break;

                case "get_link_graphics_overrides_in_view":
                    int viewIdForLink = argsObj["viewId"]?.Value<int>() ?? -1;

                    if (viewIdForLink == -1)
                    {
                        toolResult = JsonConvert.SerializeObject(new { error = "Не указан viewId" });
                        break;
                    }

                    var result_linkOverrides = Commands.GetLinkGraphicsOverridesInView(_doc, viewIdForLink);
                    toolResult = JsonConvert.SerializeObject(result_linkOverrides);
                    break;

                case "get_viewports_and_schedules_on_sheets":
                    var ids_sheets = new List<int>();
                    if (argsObj["list_elementIds"] is JArray arr_sheets)
                        ids_sheets = arr_sheets.Select(t => t.Value<int>()).ToList();

                    var result_sheets = Commands.GetViewportsAndSchedulesOnSheets(_doc, ids_sheets);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_sheets);
                    break;

                case "get_schedules_info_and_columns":
                    var ids_schedules = new List<int>();
                    if (argsObj["list_elementIds"] is JArray arr_schedules)
                        ids_schedules = arr_schedules.Select(t => t.Value<int>()).ToList();

                    var result_schedules = Commands.GetSchedulesInfoAndColumns(_doc, ids_schedules);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_schedules);
                    break;

                case "get_schedule_sorting_info":
                    var ids_sorting = new List<int>();
                    if (argsObj["list_elementIds"] is JArray arr_sorting)
                        ids_sorting = arr_sorting.Select(t => t.Value<int>()).ToList();

                    var result_sorting = Commands.GetScheduleSortingInfo(_doc, ids_sorting);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_sorting);
                    break;

                //case "get_schedule_rows_with_elements":
                //    int scheduleIdForRows = argsObj["scheduleId"]?.Value<int>() ?? -1;

                //    if (scheduleIdForRows == -1)
                //    {
                //        toolResult = JsonConvert.SerializeObject(new { error = "Не указан scheduleId", success = false });
                //        break;
                //    }

                //    var result_rows = Commands.GetScheduleRowsWithElements(_doc, scheduleIdForRows);
                //    toolResult = JsonConvert.SerializeObject(result_rows);
                //    break;


                case "get_if_elements_pass_filter":
                    int filterId = argsObj["filterId"]?.Value<int>() ?? -1;

                    var ids_passFilter = new List<int>();
                    if (argsObj["list_elementIds"] is JArray arr_passFilter)
                        ids_passFilter = arr_passFilter.Select(t => t.Value<int>()).ToList();

                    if (filterId == -1)
                    {
                        toolResult = JsonConvert.SerializeObject(new { error = "Не указан filterId" });
                        break;
                    }

                    var result_passFilter = Commands.GetIfElementsPassFilter(_doc, filterId, ids_passFilter);
                    toolResult = JsonConvert.SerializeObject(result_passFilter);
                    break;

                case "set_view_section_box_to_elements":
                    var ids_sectionBox = new List<int>();
                    if (argsObj["list_elementIds"] is JArray arr_sectionBox)
                        ids_sectionBox = arr_sectionBox.Select(t => t.Value<int>()).ToList();

                    var result_sectionBox = Commands.SetViewSectionBoxToElements(_doc, _uiDoc, ids_sectionBox);
                    toolResult = Newtonsoft.Json.JsonConvert.SerializeObject(result_sectionBox);
                    break;

                case "get_journal_entries_since":
                    string dateTimeStr = argsObj["dateTime"]?.Value<string>() ?? "";

                    if (string.IsNullOrEmpty(dateTimeStr))
                    {
                        toolResult = JsonConvert.SerializeObject(new { error = "Не указана дата", success = false });
                        break;
                    }

                    var result_journal = Commands.GetJournalEntriesSince(_doc, dateTimeStr);
                    toolResult = JsonConvert.SerializeObject(result_journal);
                    break;


                //case "get_document_switched":
                //    int linkElementId = argsObj["elementId"]?.Value<int>() ?? -1;
                //    bool switchToMain = argsObj["switchMainDoc"]?.Value<bool>() ?? false;

                //    var result_switch = Commands.GetDocumentSwitched(_doc, _uiDoc, linkElementId, switchToMain);
                //    toolResult = JsonConvert.SerializeObject(result_switch);
                //    break;




                default:
                    ChatHistory.Children.Add(CreateMessageBlock($"❓ Неизвестная команда: {toolName}", false));
                    toolResult = "[]";
                    break;
            }

            // Добавляем результат в историю для ИИ
            ChatHistoryMessages.Add(new
            {
                role = "tool",
                tool_call_id = toolCallId,
                content = toolResult
            });
        }



        private Border CreateMessageBlock(string text, bool isUser)
        {
            var border = new Border
            {
                Margin = new Thickness(5, 5, 5, 5),                                                     // Внешние отступы со всех сторон по 5px
                Padding = new Thickness(12, 8, 12, 8),                                                  // Внутренние отступы: лево/право 12, верх/низ 8
                CornerRadius = new CornerRadius(12),                                                    // Скругление углов (12px)
                Background = isUser ?                                                                   // Фон зависит от отправителя
                    new SolidColorBrush(System.Windows.Media.Color.FromRgb(59, 130, 246)) :             // Синий (#3B82F6) для пользователя
                    new SolidColorBrush(System.Windows.Media.Color.FromRgb(229, 229, 229)),             // Серый (#E5E5E5) для ИИ
                HorizontalAlignment = isUser ? HorizontalAlignment.Right : HorizontalAlignment.Left,    // Выравнивание:  
                BorderBrush = new SolidColorBrush(Colors.Gray),                                         // Цвет рамки
                BorderThickness = new Thickness(1)                                                      // Толщина рамки 1px
            };
            BindBubbleWidth(border, text, isUser);


            if (isUser)
            {
                var textBlock = new System.Windows.Controls.TextBox
                {
                    Text = text,
                    TextWrapping = TextWrapping.Wrap,
                    Foreground = isUser ? Brushes.White : Brushes.Black,
                    FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
                    FontSize = 13,
                    IsReadOnly = true,
                    Background = Brushes.Transparent,
                    BorderThickness = new Thickness(0),
                    Padding = new Thickness(0),
                    Margin = new Thickness(0),
                    VerticalContentAlignment = VerticalAlignment.Center,
                    TextAlignment = TextAlignment.Left
                };
                border.Child = textBlock;
            }
            else
            {
                var richTextBox = new RichTextBox
                {
                    IsReadOnly = true,                                                                      // Только для чтения (нельзя редактировать)
                    Background = Brushes.Transparent,                                                       // Прозрачный фон (показывается фон Border)
                    BorderThickness = new Thickness(0),                                                     // Без собственной рамки
                    Padding = new Thickness(0),                                                             // Внутренние отступы 5px
                    Margin = new Thickness(0),                                                              // Без отступов
                    FontFamily = new FontFamily("Segoe UI"),                                                // Шрифт Segoe UI
                    FontSize = 13,                                                                          // Размер шрифта 13px
                    Foreground = isUser ? Brushes.White : Brushes.Black,                                    // Цвет текста
                    VerticalScrollBarVisibility = ScrollBarVisibility.Disabled                              // Отключаем вертикальную прокрутку
                };

                // Парсим Markdown и заполняем RichTextBox
                ParseMarkdownToRichTextBox(richTextBox, text);
                border.Child = richTextBox;
            }


            //var textBlock = new System.Windows.Controls.TextBox
            //{
            //    Text = text,
            //    TextWrapping = TextWrapping.Wrap,
            //    Foreground = isUser ? Brushes.White : Brushes.Black,
            //    FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
            //    FontSize = 13,
            //    IsReadOnly = true,
            //    Background = Brushes.Transparent,
            //    BorderThickness = new Thickness(0),
            //    Padding = new Thickness(0),
            //    Margin = new Thickness(0),
            //    VerticalContentAlignment = VerticalAlignment.Center,
            //    TextAlignment = TextAlignment.Left
            //};
            //border.Child = textBlock;



            return border;

        }

        //создание интерфейса для сообщения о загрузке ИИ
        private Border CreateTypingIndicator()
        {
            var border = new Border
            {
                Margin = new Thickness(5),
                Padding = new Thickness(8),
                CornerRadius = new CornerRadius(12),
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(229, 229, 229)),
                HorizontalAlignment = HorizontalAlignment.Left,
                BorderBrush = new SolidColorBrush(Colors.Gray),
                BorderThickness = new Thickness(1)
            };
            BindBubbleWidth(border, "ИИ печатает", false);

            var richTextBox = new RichTextBox
            {
                IsReadOnly = true,
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Padding = new Thickness(5),
                Margin = new Thickness(0),
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(120, 120, 120)),
                VerticalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Height = 30,
                Width = double.NaN
            };

            var paragraph = new Paragraph();
            paragraph.Inlines.Add(new Run("ИИ печатает"));
            richTextBox.Document = new FlowDocument(paragraph);

            border.Child = richTextBox;
            return border;

        }

        private void BindBubbleWidth(Border border, string text, bool isUser)
        {
            if (border == null || ChatScrollViewer == null)
                return;

            border.SetBinding(
                FrameworkElement.WidthProperty,
                new System.Windows.Data.Binding("ActualWidth")
                {
                    Source = ChatScrollViewer,
                    Converter = new ModelBubbleWidthConverter(text, isUser),
                    ConverterParameter = "0.65"
                });
        }


        private class ModelBubbleWidthConverter : IValueConverter
        {
            private readonly string _text;
            private readonly bool _isUser;

            public ModelBubbleWidthConverter(string text, bool isUser)
            {
                _text = text ?? string.Empty;
                _isUser = isUser;
            }

            public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
            {
                double chatWidth;

                if (!(value is double) || (chatWidth = (double)value) <= 0)
                    return DependencyProperty.UnsetValue;

                double maxRatio = 0.65;

                if (parameter != null)
                {
                    double parsedRatio;
                    if (double.TryParse(
                        parameter.ToString(),
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out parsedRatio))
                    {
                        maxRatio = parsedRatio;
                    }
                }

                double minWidth = _isUser ? 140 : 180;
                double maxWidth = chatWidth * maxRatio;

                string visibleText = NormalizeBubbleText(_text);
                int longestLineLength = GetLongestLineLength(visibleText);

                double charWidth = 7.2;
                double calculatedWidth = longestLineLength * charWidth + 50;

                if (calculatedWidth < minWidth)
                    calculatedWidth = minWidth;

                if (calculatedWidth > maxWidth)
                    calculatedWidth = maxWidth;

                return calculatedWidth;
            }

            public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
            {
                throw new NotSupportedException();
            }

            private static string NormalizeBubbleText(string text)
            {
                if (string.IsNullOrWhiteSpace(text))
                    return string.Empty;

                return text
                    .Replace("\r\n", "\n")
                    .Replace("\r", "\n")
                    .Replace("**", "")
                    .Replace("__", "")
                    .Replace("`", "")
                    .Replace("###", "")
                    .Replace("##", "")
                    .Replace("#", "")
                    .Trim();
            }

            private static int GetLongestLineLength(string text)
            {
                if (string.IsNullOrWhiteSpace(text))
                    return 0;

                string[] lines = text.Split(new[] { '\n' }, StringSplitOptions.None);

                int max = 0;

                foreach (string line in lines)
                {
                    string trimmed = line == null ? string.Empty : line.Trim();

                    if (trimmed.Length > max)
                        max = trimmed.Length;
                }

                return max;
            }
        }


        private async Task<string> SendToOpenRouter(List<object> _messages)
        {
            _httpClient.DefaultRequestHeaders.Clear();
            // ⭐ ВЫБИРАЕМ URL И НАСТРОЙКИ В ЗАВИСИМОСТИ ОТ ТИПА ПОДКЛЮЧЕНИЯ ⭐
            string apiUrl;

            if (_connectionType == ConnectionType.OnlineAPI)
            {
                apiUrl = (_settings.OnlineServerUrl ?? string.Empty).Trim();
                string apiKey = (_settings.ApiKey ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(apiUrl))
                    return "Не указан URL к API внешней модели. Заполните настройку ExternalModel.OnlineServerUrl.";

                if (string.IsNullOrWhiteSpace(apiKey))
                    return "Не указан API ключ внешней модели. Заполните настройку ExternalModel.ApiKey.";

                _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");
            }
            else
            {
                apiUrl = (_settings.LocalServerUrl ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(apiUrl))
                    return "Не указан URL локальной модели. Заполните настройку ExternalModel.LocalServerUrl.";
                // Для локального сервера API-ключ не нужен
            }

            // ⭐ Устанавливаем большой таймаут для локальной модели
            if (_connectionType == ConnectionType.LocalServer)
            {
                _httpClient.Timeout = TimeSpan.FromMinutes(20);  // 20 минут
            }


            string systemPrompt = string.IsNullOrWhiteSpace(_settings.SystemPrompt)
                ? @"Ты Revit API эксперт." +
                                    "ЛОГИКА:\r\n1. Простой вопрос → ТЕКСТОВЫЙ ответ\r\n2. Техническая задача → tools + текст" +
                                    "ПРИМЕРЫ:\r\n'Расскажи о текущем виде' → ТЕКСТ: 'get_active_view_in_revit покажет детали'\r\n'Выведи текущий вид' → tool: get_active_view_in_revit() → ТЕКСТ: 'Активный вид: План 1эт'" +
                                    "НЕ возвращай RAW JSON в content. Используй tools или отвечай текстом. Если команда tools возвращает результат в футах, то пользователю выдавай результат только в миллиметрах" +
                                    "Если есть необходимость то выстраивай цепочку вызова tools, для решения задач пользователя" +
                                    "ПРИМЕР:'Покажи все окна текущего вида'" +
                                    "→ 1.get_active_view_in_revit()" +
                                    "→ 2.get_category_by_keyword('Окна')" +
                                    "→ 3.get_elements_by_category(ID_окон)"
                : _settings.SystemPrompt;

            var messagesWithSystem = new List<object> { new { role = "system", content = systemPrompt } };
            messagesWithSystem.AddRange(_messages);

            var toolsArray = new object[]
                {

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_active_view_in_revit",
                            description = "Возвращает название и ID текущего активного вида (или листа), открытого в Revit на момент вызова. Нет входных параметров.",
                            parameters = new
                            {
                                type = "object",
                                properties = new { },  // Пустой объект — нет параметров
                                required = new string[] { }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_all_elements_shown_in_view",
                            description = "Возвращает список всех element id элементов, видимых в указанном виде, на листе или в спецификации. " +
                            "Для вида — стены, колонны, помещения. Для листа — видовые экраны, основные надписи. Для спецификации — строки.",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    viewOrSheetId = new
                                    {
                                        type = "integer",
                                        description = "Element ID вида, листа или спецификации. Если не указан — текущий активный вид."
                                    }
                                },
                                required = new string[] { }
                            }
                        }

                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_category_by_keyword",
                            description = "Ищет категории Revit по ключевому слову (часть названия). Возвращает ID и имена подходящих категорий. " +
                            "Category ID нужно предварительно получить через get_category_by_keyword или get_model_categories",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    keyword = new
                                    {
                                        type = "string",
                                        description = "Ключевое слово для поиска. Примеры: 'Стены', 'Окна', 'Уровни'"
                                    }
                                },
                                required = new[] { "keyword" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_elements_by_category",
                            description = "Возвращает все element ID элементов, принадлежащих указанной категории Revit. Примеры: все стены, все двери, все уровни." +
                            "Если пользователь задал категорию именем, например Окна, то вызови комнаду get_category_by_keyword, чтобы определить id ктаегории. ",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    categoryId = new
                                    {
                                        type = "integer",
                                        description = "ID встроенной категории Revit. Примеры: -2000011=Стены, -2000240=Уровни"
                                    }
                                },
                                required = new[] { "categoryId" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_model_categories",
                            description = "Возвращает полный список всех категорий модели (системных и загружаемых). Используй ТОЛЬКО если get_category_by_keyword не нашёл нужную категорию — этот вызов может вернуть очень много данных.",
                            parameters = new
                            {
                                type = "object",
                                properties = new { },   // нет параметров
                                required = new string[] { }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_categories_from_elementids",
                            description = "Для каждого element id из списка возвращает, к какой категории он принадлежит. Удобно для определения категорий набора элементов, полученных из других инструментов.",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    list_elementIds = new
                                    {
                                        type = "array",
                                        items = new { type = "integer" },
                                        description = "Список element id, для которых нужно определить категории"
                                    }
                                },
                                required = new[] { "list_elementIds" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_element_types_for_elementids",
                            description = "Для каждого element id возвращает его type id и имя типа. Позволяет узнать, к какому типу принадлежит каждый элемент — например, тип стены '200мм Кирпич' или тип двери 'ДВ-1'.",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    list_elementIds = new
                                    {
                                        type = "array",
                                        items = new { type = "integer" },
                                        description = "Список element id для определения типов"
                                    }
                                },
                                required = new[] { "list_elementIds" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_all_elementids_for_specific_type_ids",
                            description = "Обратная операция к get_element_types_for_elementids. По type id возвращает все экземпляры (element id) данного типа в модели. Полезно для поиска всех вхождений конкретного типа.",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    list_typeIds = new  // ← typeIds, не elementIds!
                                    {
                                        type = "array",
                                        items = new { type = "integer" },
                                        description = "Список TYPE ID (не element id!). Получи через get_element_types_for_elementids."
                                    }
                                },
                                required = new[] { "list_typeIds" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_all_used_families_in_model",
                            description = "Возвращает все семейства в модели Revit (как загружаемые, так и системные). " +
                                          "Для каждого семейства возвращает: " +
                                          "- FamilyId: уникальный идентификатор (для системных — отрицательный хэш имени) " +
                                          "- FamilyName: имя семейства " +
                                          "- IsLoadedFamily: true = загружаемое (дверь, окно, мебель), false = системное (стена, перекрытие, уровень) " +
                                          "- IsPlacedInModel: true = размещено в модели (есть хотя бы один экземпляр) " +
                                          "- InstanceCount: количество экземпляров в модели " +
                                          "Дополнительно возвращает stats с общей статистикой: " +
                                          "- loaded_families: количество загружаемых семейств " +
                                          "- system_families: количество системных семейств " +
                                          "- placed_families: количество размещённых семейств " +
                                          "- unplaced_families: количество загруженных, но не размещённых (можно удалить) " +
                                          "Полезно для: аудита семейств, поиска неиспользуемых элементов, анализа состава проекта.",
                            parameters = new
                            {
                                type = "object",
                                properties = new { },
                                required = new string[] { }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_all_used_families_of_category",
                            description = "Возвращает все загружаемые семейства конкретной категории. Аналогичен get_all_used_families_in_model (все те же праарметры возвращает), но возвращает только семейства нужной категории.",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    categoryId = new
                                    {
                                        type = "integer",
                                        description = "ID категории Revit (получить через get_category_by_keyword)."
                                    }
                                },
                                required = new[] { "categoryId" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_all_used_types_of_a_family",
                            description = "По точному имени семейства возвращает все его типоразмеры (type). Работает как с загружаемыми, так и с системными семействами (стены, перекрытия). Имя должно совпадать точно.",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    familyName = new
                                    {
                                        type = "string",
                                        description = "Точное имя семейства. Пример: 'Базовая стена'. Чувствительно к регистру."
                                    }
                                },
                                required = new[] { "familyName" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_all_elements_of_specific_families",
                            description = "По списку точных имён семейств возвращает все element id экземпляров этих семейств в модели. Позволяет найти все вхождения нескольких семейств за один вызов.",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    familyNames = new
                                    {
                                        type = "array",
                                        items = new { type = "string" },
                                        description = "Список точных имён семейств. Пример: ['Базовая стена','Дверь ДВ-1']"
                                    }
                                },
                                required = new[] { "familyNames" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_parameters_from_elementid",
                            description = "Возвращает ВСЕ параметры (по каждому парметру это: id параметра, имя, значение, storageType - типзначения, isReadOnly - только ли на чтение параметр (true) или нет(false)) одного конкретного элемента. Это основной инструмент для изучения доступных параметров. Рекомендуется вызывать первым перед массовым get_parameter_value_for_element_ids — чтобы узнать нужный idParameter. Обрати внимание что единица измерения в Revit не метры или миллиметры, а футы",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    elementId = new
                                    {
                                        type = "integer",
                                        description = "Element id одного элемента или типа."
                                    },
                                    getIdValuesAsNames = new
                                    {
                                        type = "boolean",
                                        description = "Если true — ElementId параметры возвращаются как имена связанных элементов; если false — как числовые ID."
                                    }
                                },
                                required = new[] { "elementId", "getIdValuesAsNames" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_parameter_value_for_element_ids",
                            description = "Получает значение одного конкретного параметра для большого списка элементов. Используется для массового извлечения данных после того, как нужный parameterId найден через get_parameters_from_elementid. Обрати внимание что единица измерения в Revit не метры или миллиметры, а футы",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    list_elementIds = new
                                    {
                                        type = "array",
                                        items = new { type = "integer" },
                                        description = "Список element id, для которых нужно получить значение параметра."
                                    },
                                    idParameter = new
                                    {
                                        type = "integer",
                                        description = "ID параметра (получить через get_parameters_from_elementid)."
                                    },
                                    getIdValuesAsNames = new
                                    {
                                        type = "boolean",
                                        description = "Если true — ElementId параметры возвращаются как имена связанных элементов; если false — как числовые ID."
                                    }
                                },
                                required = new[] { "list_elementIds", "idParameter", "getIdValuesAsNames" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_all_additional_properties_from_elementid",
                            description = "Возвращает дополнительные свойства одного элемента, доступные через Revit API классы (не через параметры). Используйте только если get_parameters_from_elementid не вернул нужные данные. Возвращает имена и значения свойств без их id.",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    elementId = new
                                    {
                                        type = "integer",
                                        description = "Element id одного элемента."
                                    }
                                },
                                required = new[] { "elementId" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_additional_property_for_all_elementids",
                            description = "Массовая версия get_all_additional_properties_from_elementid — получает одно конкретное дополнительное свойство (по имени) доступное через Revit API классы (не через параметры) для списка элементов. Имя свойства должно совпадать точно.",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    list_elementIds = new
                                    {
                                        type = "array",
                                        items = new { type = "integer" },
                                        description = "Список element id."
                                    },
                                    propertyName = new
                                    {
                                        type = "string",
                                        description = "Точное имя свойства (как в get_all_additional_properties_from_elementid)."
                                    }
                                },
                                required = new[] { "list_elementIds", "propertyName" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_location_for_element_ids",
                            description = "Возвращает точку или кривую расположения для списка элементов. Для точечных объектов (колонны, двери) — координаты XYZ точки. Для линейных (стены, трубы) — начальная и конечная точки кривой.",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    list_elementIds = new
                                    {
                                        type = "array",
                                        items = new { type = "integer" },
                                        description = "Список element id."
                                    }
                                },
                                required = new[] { "list_elementIds" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_boundingboxes_for_element_ids",
                            description = "Возвращает ограничивающий прямоугольник (BoundingBoxXYZ) для списка элементов. Минимальные и максимальные координаты XYZ. Габариты + расположение. Координаты в футах.",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    list_elementIds = new
                                    {
                                        type = "array",
                                        items = new { type = "integer" },
                                        description = "Список Element ID"
                                    },
                                    idSheet = new
                                    {
                                        type = "integer",
                                        description = "ID листа (опционально, для аннотаций/видов)"
                                    }
                                },
                                required = new[] { "list_elementIds" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_boundary_lines",
                            description = @"Возвращает точные граничные линии (рёбра геометрии) для элементов Revit.

                            ВОЗВРАЩАЕМЫЕ ДАННЫЕ:
                            Для каждого element id возвращается:
                            - ElementId: ID элемента
                            - Lines: массив линий, каждая линия содержит:
                                • StartX, StartY, StartZ - координаты начала линии (в футах)
                                • EndX, EndY, EndZ - координаты конца линии (в футах)
                                • Length - длина линии (в футах)
                            - LineCount: количество найденных линий
                            - Error: сообщение об ошибке (если есть)

                            ОСОБЕННОСТИ:
                            1. Координаты возвращаются в футах (1 фут = 304.8 мм)
                            2. Для перевода в миллиметры умножьте на 304.8
                            3. Для помещений используются границы, определённые через инструмент 'Границы помещения'
                            4. Для криволинейных элементов (арки, дуги) линии аппроксимируются прямыми отрезками
                            5. Возвращает ВСЮ геометрию элемента, включая вложенные семейства

                            ПРИМЕРЫ ИСПОЛЬЗОВАНИЯ:
                            - 'Покажи все рёбра выбранной стены'
                            - 'Найди длину нижнего ребра перекрытия'
                            - 'Определи границы комнаты 101'
                            - 'Проверь пересекаются ли эти две балки'
                            - 'Получи форму колонны'

                            ПРИМЕЧАНИЕ:
                            - Не все элементы имеют геометрию (уровни, сетки, текстовые заметки вернут пустой результат)
                            - Для сложной геометрии количество линий может быть большим (тысячи)
                            - Для помещений рекомендуется использовать отдельный метод get_room_boundary_lines, если нужны только границы помещения",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    list_elementIds = new
                                    {
                                        type = "array",
                                        items = new { type = "integer" },
                                        description = "Список element id элементов. Поддерживаемые типы: стены (Walls), перекрытия (Floors), " +
                                                        "колонны (Columns), балки (Beams), крыши (Roofs), фундаменты (Footings), " +
                                                        "помещения (Rooms), семейства (FamilyInstances) - мебель, оборудование, сантехника, " +
                                                        "а также любые другие элементы с 3D-геометрией."
                                    }
                                },
                                required = new[] { "list_elementIds" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_room_boundary_lines",
                            description = @"Специализированная команда для получения границ помещений Revit.

                            ВОЗВРАЩАЕМЫЕ ДАННЫЕ:
                            Для каждого помещения возвращается:
                            - ElementId: ID помещения
                            - Lines: массив линий границ, каждая линия содержит:
                              • StartX, StartY, StartZ - координаты начала линии (в футах)
                              • EndX, EndY, EndZ - координаты конца линии (в футах)
                              • Length - длина линии (в футах)
                            - LineCount: количество найденных линий
                            - Error: сообщение об ошибке (если есть)

                            ОСОБЕННОСТИ РАБОТЫ С ПОМЕЩЕНИЯМИ:
                            1. Границы рассчитываются на основе настроек 'Границы помещения' в Revit:
                               - Стены (включая многослойные)
                               - Колонны
                               - Перегородки
                               - Ограждения
                               - Виртуальные границы
                            2. Учитываются вырезы в помещениях (колонны, шахты, ниши)
                            3. Для помещений без границ (незамкнутый контур) возвращается ошибка
                            4. Координаты возвращаются в футах (1 фут = 304.8 мм)
                            5. Для перевода в миллиметры умножьте на 304.8
                            6. Криволинейные границы (дуги) аппроксимируются прямыми отрезками

                            ОТЛИЧИЯ ОТ get_boundary_lines:
                            - get_boundary_lines: возвращает ВСЮ геометрию элемента (все рёбра 3D-тела)
                            - get_room_boundary_lines: возвращает ТОЛЬКО границы помещения (2D-контур на уровне пола)
                            - get_room_boundary_lines автоматически обрабатывает соединения между стенами
                            - get_room_boundary_lines учитывает правила расчёта площади помещения Revit

                            ПРИМЕРЫ ИСПОЛЬЗОВАНИЯ:
                            - 'Покажи план помещения 101'
                            - 'Рассчитай периметр комнаты Конференц-зал'
                            - 'Найди все помещения с площадью больше 50 кв.м'
                            - 'Проверь, какие комнаты граничат с коридором'
                            - 'Построй 3D-модель планировки этажа'
                            - 'Найди помещения неправильной формы'
                            - 'Определи соседние помещения по общим границам'

                            РАСЧЁТЫ НА ОСНОВЕ ГРАНИЦ:
                            - Периметр помещения = сумма длин всех линий
                            - Площадь помещения (можно также использовать стандартное свойство Room.Area)
                            - Форма помещения (прямоугольное/непрямоугольное)
                            - Количество углов помещения
                            - Максимальные/минимальные размеры

                            ПРИМЕЧАНИЯ:
                            - Если помещение не размечено (Unplaced Room), команда вернёт ошибку
                            - Для помещений на разных уровнях (Level) координата Z будет соответствовать высоте уровня
                            - Для помещений с вырезами границы возвращаются с учётом всех отверстий
                            - При использовании с большим количеством помещений может потребоваться время",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    list_roomIds = new
                                    {
                                        type = "array",
                                        items = new { type = "integer" },
                                        description = "Список ID помещений. Примеры получения ID помещений: " +
                                                     "1. Через get_elements_by_category с ID категории -2000050 (категория Rooms) " +
                                                     "2. Через get_all_elements_shown_in_view для видов с помещениями " +
                                                     "3. Через get_room_boundary_lines без параметров (вернёт все помещения) " +
                                                     "Если параметр не указан или передан пустой массив, возвращаются все помещения модели. " +
                                                     "Пример: list_roomIds = [123456, 789012, 345678]"
                                    }
                                },
                                required = new string[] { }  // Необязательный параметр - можно вызывать без ID для получения всех помещений
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_host_id_for_element_ids",
                            description = "Для размещённых элементов (окна, двери, сантехника) возвращает ID хост-элемента (стены, перекрытия).",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    list_elementIds = new
                                    {
                                        type = "array",
                                        items = new { type = "integer" },
                                        description = "Список ID размещённых элементов (окна, двери)"
                                    }
                                },
                                required = new[] { "list_elementIds" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_object_classes_from_elementids",
                            description = "Возвращает полное имя C#-класса Revit API для каждого элемента. Позволяет узнать программный тип объекта.",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    list_elementIds = new
                                    {
                                        type = "array",
                                        items = new { type = "integer" },
                                        description = "Список element id"
                                    }
                                },
                                required = new[] { "list_elementIds" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_material_layers_from_types",
                            description = "Для системных типов (WallType, FloorType, RoofType, CeilingType) возвращает слои конструкции: материал, толщину (в футах) и функцию слоя. Вход — TYPE id, не element id.",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    list_elementIds = new
                                    {
                                        type = "array",
                                        items = new { type = "integer" },
                                        description = "Список TYPE id системных конструкций (WallType, FloorType, RoofType, CeilingType)."
                                    }
                                },
                                required = new[] { "list_elementIds" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_model_file_info",
                            description = "Возвращает информацию о файле текущей модели Revit: путь расположения и размер в МБ. " +
                                          "Если открыта локальная копия центрального файла (workshared), дополнительно возвращает " +
                                          "путь и размер центрального файла-хранилища. Полезно для мониторинга размеров модели, " +
                                          "поиска путей к файлам, проверки использования дискового пространства.",
                            parameters = new
                            {
                                type = "object",
                                properties = new { },  // Нет входных параметров
                                required = new string[] { }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_all_project_units",
                            description = "Возвращает все единицы измерения, настроенные в текущем проекте Revit. " +
                                          "Включает типы: длина, площадь, объём, угол, масса, температура, стоимость и другие. " +
                                          "Помогает правильно интерпретировать числовые значения параметров, так как Revit может " +
                                          "использовать разные системы единиц (метрическую или имперскую). " +
                                          "Возвращает для каждого типа: название типа, символ единицы и тип отображения. " +
                                          "Используйте эту команду перед анализом числовых параметров, чтобы понять, " +
                                          "в каких единицах получены значения (например, при получении длины стены или площади помещения).",
                            parameters = new
                            {
                                type = "object",
                                properties = new { },
                                required = new string[] { }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_all_warnings_in_the_model",
                            description = "Возвращает все предупреждения модели Revit. " +
                                          "Анализирует документ и собирает информацию о проблемах: пересекающиеся стены, " +
                                          "незамкнутые помещения, несоединенные элементы и т.д. " +
                                          "Помогает выявлять и исправлять проблемы модели, повышая её качество. " +
                                          "Возвращает для каждого предупреждения: описание (текст), серьезность (всегда 'Warning') " +
                                          "и список ID элементов, связанных с этим предупреждением. " +
                                          "Используйте эту команду для аудита модели перед экспортом или проверки целостности данных.",
                            parameters = new
                            {
                                type = "object",
                                properties = new { },
                                required = new string[] { }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_all_workset_information",
                            description = "Возвращает полную информацию о всех рабочих наборах (worksets) в текущем проекте Revit. " +
                                          "Рабочие наборы используются для разделения модели на логические части при совместной работе (worksharing). " +
                                          "Команда возвращает список всех наборов с детальной информацией о каждом: " +
                                          "ID набора, имя, владелец (кто сейчас редактирует), редактируемость текущим пользователем, " +
                                          "тип набора (пользовательский/семейства/виды/стандарты), статус открытия, является ли набором по умолчанию. " +
                                          "Важно: команда работает только для общих (workshared) документов. Если документ не является общим, " +
                                          "будет возвращена соответствующая ошибка. " +
                                          "Поле isEditable показывает, может ли текущий пользователь редактировать элементы этого набора. " +
                                          "Используйте эту команду для анализа текущего состояния рабочих наборов, проверки прав доступа, " +
                                          "планирования совместной работы или аудита проекта.",
                            parameters = new
                            {
                                type = "object",
                                properties = new { },
                                required = new string[] { }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_worksets_from_elementids",
                            description = "Для каждого element id возвращает информацию о рабочем наборе (workset), в котором находится элемент. " +
                                          "Позволяет быстро проверить принадлежность элементов к рабочим наборам и их редактируемость. " +
                                          "Возвращает для каждого элемента: ID рабочего набора, имя рабочего набора, может ли текущий пользователь " +
                                          "редактировать этот набор (isEditable), тип (пользовательский/семейства/виды/стандарты) и открыт ли рабочий набор. " +
                                          "Полезно для: проверки прав доступа к элементам перед модификацией, анализа распределения элементов по рабочим наборам, " +
                                          "выявления проблем с редактированием, аудита совместной работы. " +
                                          "Важно: команда работает только для общих (workshared) документов. Для необщих документов будет возвращена ошибка. " +
                                          "Если элемент не принадлежит рабочему набору (например, системный элемент), это будет указано в результате.",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    list_elementIds = new
                                    {
                                        type = "array",
                                        items = new { type = "integer" },
                                        description = "Список element id элементов, для которых нужно определить рабочий набор. " +
                                                     "Поддерживаются любые элементы Revit: стены, двери, семейства, виды и т.д."
                                    }
                                },
                                required = new[] { "list_elementIds" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_worksharing_information_for_element_ids",
                            description = "Расширенная версия get_worksets_from_elementids. Возвращает полную информацию о совместной работе для указанных элементов Revit: " +
                                          "рабочий набор элемента, создатель элемента, текущий владелец, кто последний изменял элемент, статус редактирования. " +
                                          "Использует WorksharingUtils.GetWorksharingTooltipInfo() для получения детальной информации [citation:4][citation:5]. " +
                                          "Полезно для: аудита изменений в модели, выявления авторов элементов, проверки прав доступа, " +
                                          "анализа истории изменений, решения конфликтов при совместной работе. " +
                                          "Важно: команда работает только для общих (workshared) документов. " +
                                          "Информация основана на локальном кэше и может быть немного устаревшей — это нормально для целей отображения пользователю [citation:2].",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    list_elementIds = new
                                    {
                                        type = "array",
                                        items = new { type = "integer" },
                                        description = "Список element id элементов для анализа"
                                    }
                                },
                                required = new[] { "list_elementIds" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_user_selection_in_revit",
                            description = "Возвращает element id элементов, которые пользователь в данный момент выделил в Revit. " +
                                          "Позволяет ИИ работать именно с теми элементами, на которые пользователь обратил внимание. " +
                                          "Возвращает список ID выделенных элементов, а также базовую информацию о каждом (имя, категория). " +
                                          "Полезно для: анализа выбранных элементов, выполнения операций с выделением, " +
                                          "получения контекста для дальнейших запросов. " +
                                          "Важно: если ничего не выделено, возвращается пустой список с соответствующим сообщением.",
                            parameters = new
                            {
                                type = "object",
                                properties = new { },
                                required = new string[] { }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "set_user_selection_in_revit",
                            description = "Устанавливает выделение элементов в Revit. Пользователь увидит подсвеченные элементы. " +
                                          "ВАЖНО: ПЕРЕЗАПИСЫВАЕТ текущее выделение пользователя. " +
                                          "Вызывать не более одного раза за рабочий процесс (после того как найдены нужные элементы). " +
                                          "ТОЛЬКО element id экземпляров (не type id, не семейства, не категории). " +
                                          "Документ должен быть тем же, из которого получены ID. " +
                                          "Полезно для: визуальной индикации найденных элементов, подсветки проблемных элементов, " +
                                          "помощи пользователю в навигации по модели. " +
                                          "Возвращает success (успех) и количество выделенных элементов. " +
                                          "Если переданные ID не являются element id экземпляров или не существуют, они будут пропущены с пояснением в ответе.",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    list_elementIds = new
                                    {
                                        type = "array",
                                        items = new { type = "integer" },
                                        description = "Список element id экземпляров для выделения. " +
                                                     "Только ID экземпляров (не типов, не семейств, не категорий). " +
                                                     "Пример: [500, 501, 502] - выделит 3 элемента. " +
                                                     "Если список пуст, выделение будет сброшено."
                                    }
                                },
                                required = new[] { "list_elementIds" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_graphic_overrides_for_element_ids_in_view",
                            description = "Возвращает графические переопределения конкретных элементов в указанном виде Revit. " +
                                          "Переопределения имеют высший приоритет графики (переопределяют Object Styles, категории, фильтры). " +
                                          "Возвращает информацию о цветах линий проекции и разреза, паттернах заливки, видимости, полутоне, прозрачности. " +
                                          "ВАЖНО: НЕ работает для элементов из связанных документов (Linked Files). " +
                                          "ТОЛЬКО для element id экземпляров (не type id!). " +
                                          "Полезно для: аудита настроек видимости, проверки переопределений перед экспортом, " +
                                          "выявления элементов с нестандартной графикой. " +
                                          "Возвращает для каждого элемента: projection (линии проекции/поверхности), cut (линии разреза), " +
                                          "is_hidden (скрыт ли), halftone (полутон), has_overrides (есть ли активные переопределения).",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    list_elementIds = new
                                    {
                                        type = "array",
                                        items = new { type = "integer" },
                                        description = "Список element id экземпляров элементов для проверки переопределений"
                                    },
                                    viewId = new
                                    {
                                        type = "integer",
                                        description = "Element id вида, в котором проверяются переопределения"
                                    }
                                },
                                required = new[] { "list_elementIds", "viewId" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_graphic_filters_applied_to_views",
                            description = "Возвращает все фильтры видов, применённые к указанным видам Revit. " +
                                          "Для каждого фильтра возвращает: " +
                                          "- filterId: ID фильтра " +
                                          "- filterName: имя фильтра " +
                                          "- categories: список ID категорий, к которым применяется фильтр " +
                                          "- isFilterVisible: (bool) видимость элементов, прошедших фильтр (столбец 'Видимость'). Доступно для всех версий Revit 2014+. " +
                                          "- hasRules: есть ли у фильтра правила " +
                                          "- ruleParameters: список ID параметров в правилах " +
                                          "- revitVersion: версия Revit " +
                                          "ВАЖНОЕ ОГРАНИЧЕНИЕ API: " +
                                          "Статус 'Включен/Выключен' фильтра (галочка в столбце 'Включить') НЕДОСТУПЕН через Revit API в версиях ниже 2025. " +
                                          "Если пользователь спрашивает 'включён ли фильтр?' или 'активен ли фильтр?', вы должны ответить: " +
                                          "\"В связи с ограничением Revit API я не могу получить информацию о том, включён ли этот фильтр на виде. " +
                                          "Вы можете самостоятельно проверить это в настройках вида: Видимость/Графика → вкладка 'Фильтры' → столбец 'Включить фильтр'.\" " +
                                          "НЕ ПЫТАЙТЕСЬ определить статус 'Включено' по isFilterVisible или другим полям — это разные понятия. " +
                                          "Полезно для: аудита настроек фильтров, проверки видимости элементов, анализа применённых фильтров.",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    list_elementIds = new
                                    {
                                        type = "array",
                                        items = new { type = "integer" },
                                        description = "Список element id видов или листов Revit"
                                    }
                                },
                                required = new[] { "list_elementIds" }
                            }
                        }
                    },


                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_all_parameter_filters_in_model",
                            description = "Возвращает все фильтры видов (ParameterFilterElement) в текущей модели Revit. " +
                                            "Фильтры видов используются для переопределения графики элементов на видах на основе их параметров. " +
                                            "Выходные данные: " +
                                            "- filterId: ID фильтра " +
                                            "- filterName: имя фильтра " +
                                            "- categories: список ID категорий, к которым применяется фильтр " +
                                            "- hasRules: есть ли у фильтра правила фильтрации " +
                                            "- ruleParameters: список ID параметров, участвующих в правилах фильтра " +
                                            "Полезно для: аудита всех фильтров в проекте, поиска фильтров по категориям, " +
                                            "понимания структуры фильтрации, подготовки данных для get_if_elements_pass_filter. " +
                                            "Важно: команда не принимает параметров и возвращает все фильтры в документе.",
                            parameters = new
                            {
                                type = "object",
                                properties = new { },
                                required = new string[] { }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_graphic_overrides_view_filters",
                            description = "Возвращает графические настройки (переопределения) для конкретных фильтров в указанном виде Revit. " +
                                          "Фильтры имеют средний приоритет графики — могут быть переопределены поэлементными настройками, " +
                                          "но переопределяют настройки категорий и Object Styles. " +
                                          "НЕ работает для элементов из связанных документов. " +
                                          "Возвращает для каждого фильтра: " +
                                          "- projection: настройки линий проекции (цвет, образец, вес), штриховка передней поверхности (образец, цвет, видимость (если значение false значит видимость выключена))," +
                                          "штриховка задней поверхности (образец, цвет, видимость (если значение false значит видимость выключена)), прозрачность " +
                                          "- cut: настройки линий сечения (цвет, образец, вес), штриховка переднего сечения (образец, цвет, видимость (если значение false значит видимость выключена))," +
                                          "штриховка заднего сечения (образец, цвет, видимость (если значение false значит видимость выключена)) " +
                                          "- halftone: полутон " +
                                          "- has_overrides: есть ли активные переопределения " +
                                          "Полезно для: аудита настроек фильтров, понимания графики в виде, диагностики конфликтов переопределений. " +
                                          "Важно: фильтр должен быть применён к виду (иначе вернётся ошибка). " +
                                          "Для получения списка фильтров, применённых к виду, используйте get_graphic_filters_applied_to_views." +
                                          "ВАЖНО: При ответе пользователю ОБЯЗАТЕЛЬНО детально описывайте каждый параметр: " +
                                          "- Для цвета штриховки указывайте значения RGB (красный, зеленый, синий) " +
                                          "- Для видимости штриховки указывайте 'включена' (true) или 'выключена' (false) " +
                                          "- Если параметр отсутствует (null), сообщайте, что переопределение не задано. " +
                                          "НЕ сокращайте ответ. Выводите информацию о цвете и видимости для каждой штриховки.",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    list_filterIds = new
                                    {
                                        type = "array",
                                        items = new { type = "integer" },
                                        description = "Список filter id фильтров, для которых нужно получить графические переопределения"
                                    },
                                    viewId = new
                                    {
                                        type = "integer",
                                        description = "Element id вида, в котором применены фильтры"
                                    }
                                },
                                required = new[] { "list_filterIds", "viewId" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_category_visibility_overrides_in_view",
                            description = "Возвращает переопределения видимости по категориям на указанном виде Revit. " +
                                          "Позволяет получить информацию о том, какие категории скрыты на виде, а также их графические переопределения " +
                                          "(цвета линий, паттерны заливки, вес линий, прозрачность и т.д.). " +
                                          "Выходные данные: " +
                                          "- categories_overrides: словарь, где ключ - ID категории, значение - информация о категории: " +
                                          "    • category_id: ID категории " +
                                          "    • category_name: имя категории (например, 'Стены', 'Двери', 'Окна') " +
                                          "    • is_hidden: скрыта ли категория на виде (true/false) " +
                                          "    • overrides: детальные настройки переопределений: " +
                                          "        - projection: настройки проекции (линии, заливка поверхности, прозрачность) " +
                                          "        - cut: настройки разреза (линии, заливка) " +
                                          "        - halftone: полутон " +
                                          "        - has_overrides: наличие активных переопределений (детальная разбивка по типам) " +
                                          "- count: количество обработанных категорий " +
                                          "- processed_successfully: количество успешно обработанных категорий " +
                                          "- view_id: ID вида " +
                                          "- view_name: имя вида " +
                                          "- view_type: тип вида " +
                                          "Поддерживаемые типы видов: планы этажей, планы потолков, 3D-виды, чертёжные виды, фасады, разрезы. " +
                                          "Полезно для: аудита настроек видимости, диагностики проблем с отображением элементов, " +
                                          "проверки, почему определённые категории не видны на виде.",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    viewId = new
                                    {
                                        type = "integer",
                                        description = "Element id вида, для которого нужно получить переопределения категорий"
                                    }
                                },
                                required = new[] { "viewId" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_workset_visibility_in_view",
                            description = "Возвращает информацию о видимости рабочих наборов на указанном виде Revit. " +
                                          "Работает только для общих (workshared) документов. " +
                                          "Выходные данные: " +
                                          "- view_id: ID вида " +
                                          "- view_name: имя вида " +
                                          "- view_type: тип вида " +
                                          "- workset_visibility: список рабочих наборов с информацией о видимости " +
                                          "- count: количество рабочих наборов " +
                                          "Для каждого рабочего набора возвращается: " +
                                          "- workset_id: ID рабочего набора " +
                                          "- workset_name: имя рабочего набора " +
                                          "- visibility_status: статус переопределения на виде (Visible/Hidden/UseGlobalSetting) " +
                                          "- visibility_status_ru: статус на русском (показать/скрыть/использовать глобальную настройку видимости (видимый/невидимый)) " +
                                          "- globally_visible: глобальная настройка видимости рабочего набора (IsVisibleByDefault) " +
                                          "Полезно для: проверки настроек видимости рабочих наборов на видах, диагностики проблем с отображением элементов, " +
                                          "аудита совместной работы.",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    viewId = new
                                    {
                                        type = "integer",
                                        description = "Element id вида, для которого нужно получить информацию о видимости рабочих наборов"
                                    }
                                },
                                required = new[] { "viewId" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_link_graphics_overrides_in_view",
                            description = "Возвращает информацию о графических переопределениях связанных файлов (RevitLinkInstance) на указанном виде Revit. " +
                                          "Команда доступна ТОЛЬКО для Revit 2024 и выше. Для более старых версий будет возвращено сообщение о недоступности. " +
                                          "Позволяет определить настройки видимости, полутона и режима отображения для каждого связанного файла на виде. " +
                                          "Выходные данные: " +
                                          "- success: успешность выполнения " +
                                          "- revit_version: версия Revit (для проверки совместимости) " +
                                          "- view_id: ID вида " +
                                          "- view_name: имя вида " +
                                          "- view_type: тип вида " +
                                          "- link_overrides: список связанных файлов с информацией о переопределениях " +
                                          "- count: количество связанных файлов " +
                                          "Для каждого связанного файла возвращается: " +
                                          "- link_instance_id: ID экземпляра связанного файла " +
                                          "- link_type_id: ID типа связанного файла (RevitLinkType) " +
                                          "- link_name: имя связанного файла " +
                                          "- is_hidden: скрыт ли связанный файл на виде (true/false) " +
                                          "- is_halftone: включён ли режим полутона (true/false) " +
                                          "- display_setting: режим отображения на русском (по основному виду/по связанному виду/пользовательский) " +
                                          "- linked_view: информация о связанном виде (если display_setting = по связанному виду): " +
                                          "    • view_id: ID вида " +
                                          "    • view_name: имя вида " +
                                          "    • view_type: тип вида " +
                                          "Важно: команда работает только для Revit 2024 и выше. " +
                                          "Поддерживаемые типы видов: планы этажей, планы потолков, 3D-виды, чертёжные виды, фасады, разрезы. " +
                                          "Полезно для: аудита настроек связанных файлов, диагностики проблем с отображением, " +
                                          "проверки, почему связанные элементы не видны на виде.",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    viewId = new
                                    {
                                        type = "integer",
                                        description = "Element id вида, для которого нужно получить переопределения связанных файлов"
                                    }
                                },
                                required = new[] { "viewId" }
                            }
                        }
                    },


                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_viewports_and_schedules_on_sheets",
                            description = "Возвращает все объекты, размещённые на указанных листах Revit: видовые экраны (viewport), спецификации (schedules), а также другие элементы: текст, размеры, штампы, линии детализации, облака ревизий, изображения, семейства. " +
                                          "Позволяет получить полную структуру листа и понять, какие объекты на нём расположены. " +
                                          "Выходные данные для каждого листа: " +
                                          "- sheet_name: имя листа " +
                                          "- sheet_number: номер листа " +
                                          "- contents: список всех объектов на листе (единый массив) " +
                                          "- count: общее количество объектов " +
                                          "- viewports_count: количество видовых экранов " +
                                          "- schedules_count: количество спецификаций " +
                                          "- other_elements_count: количество остальных элементов " +
                                          "Для КАЖДОГО ОБЪЕКТА (в массиве contents) возвращается: " +
                                          "- viewportId: ID объекта (видового экрана, спецификации или элемента) " +
                                          "- referencedViewId: ID ссылочного вида (только для видовых экранов и спецификаций, для остальных элементов = null) " +
                                          "- viewName: имя вида/спецификации или содержимое элемента (текст, размер и т.д.). У текстовых примечаний поле viewName содержит текст данного примечания" +
                                          "- type: тип объекта на русском языке (План этажа, Разрез, 3D вид, Спецификация, Текст, Размер, Основная надпись (штамп), Линия детализации, Облако ревизии, Изображение, Семейство и т.д.) " +
                                          "Важно: команда возвращает ВСЕ объекты на листе в одном массиве contents, включая виды, спецификации, текст, размеры, штампы и другие элементы. " +
                                          "Это позволяет получить полную картину содержимого листа для анализа и аудита документации. " +
                                          "Полезно для: анализа состава листов, проверки правильности размещения видов и элементов, аудита документации, " +
                                          "поиска текстовых пометок, проверки наличия штампов, выявления пустых листов.",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    list_elementIds = new
                                    {
                                        type = "array",
                                        items = new { type = "integer" },
                                        description = "Список element id листов (ViewSheet) для анализа"
                                    }
                                },
                                required = new[] { "list_elementIds" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_schedules_info_and_columns",
                            description = @"Анализирует спецификацию Revit и возвращает полную информацию о её структуре: столбцы, фильтры, параметры.

                    ВОЗВРАЩАЕМЫЕ ДАННЫЕ:

                    Для каждой спецификации (scheduleId):
                    - scheduleId: ID спецификации (ViewSchedule)
                    - scheduleName: имя спецификации
                    - categoryId: ID категории (например, -2000011 для стен)
                    - categoryName: имя категории (например, 'Стены')
                    - rowCount: количество строк данных в спецификации
                    - columns: список столбцов
                    - columnsCount: количество столбцов
                    - filters: список фильтров
                    - filtersCount: количество фильтров
                    - hasFilters: наличие активных фильтров

                    Для КАЖДОГО СТОЛБЦА (columns):
                    - header: заголовок столбца (что видит пользователь)
                    - parameterName: техническое имя параметра
                    - parameterId: ID параметра (-1 для вычисляемых полей)
                    - isHidden: скрыт ли столбец
                    - isCalculated: является ли вычисляемым (формула/процент/количество)
                    - isCombinedParameter: объединённый ли параметр (несколько полей в одном)
                    - fieldType: тип поля (Formula/Percentage/Count/CombinedParameter и т.д.)
                    - fieldTypeDescription: понятное описание типа поля
                    - calculatedType: подтип вычисляемого поля (Formula/Percentage/Count)
                    - percentageOfField: для поля Percentage — имя поля, от которого считается процент
                    - percentageByField: для поля Percentage — имя поля для группировки
                    - combinedParameters: для CombinedParameter — список объединяемых параметров:
                        • parameterId: ID параметра
                        • parameterName: имя параметра
                        • prefix: префикс перед значением
                        • separator: разделитель между параметрами
                        • suffix: суффикс после значения
                        • sample: образец отображения

                    Для КАЖДОГО ФИЛЬТРА (filters):
                    - fieldId: ID поля, к которому применён фильтр
                    - fieldName: имя поля
                    - fieldParameterId: ID параметра (если доступен)
                    - filterType: тип сравнения (Equal/GreaterThan/Contains и т.д.)
                    - filterTypeDescription: понятное описание (равно/больше/содержит)
                    - value: значение фильтра (строка/число/объект с id,name)
                    - valueType: тип значения (String/Integer/Double/ElementId/Null)

                    ПРАВИЛА:
                    - Фильтры работают по принципу 'И' (AND) — элемент должен удовлетворять ВСЕМ условиям.
                    - isCalculated = true → parameterId = -1, значения таких полей уже есть в get_all_elements_shown_in_view.
                    - isCombinedParameter = true → значения формируются из нескольких параметров.

                    АЛГОРИТМ РАБОТЫ СО СПЕЦИФИКАЦИЕЙ:
                    1. Вызвать get_schedules_info_and_columns — получить структуру
                    2. Вызвать get_all_elements_shown_in_view с ID спецификации — получить строки
                    3. Для parameterId != -1: get_parameter_value_for_element_ids
                    4. Для parameterId = -1: значения уже есть в get_all_elements_shown_in_view (поле 'name')",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    list_elementIds = new
                                    {
                                        type = "array",
                                        items = new { type = "integer" },
                                        description = "Список element id спецификаций (ViewSchedule) для анализа"
                                    }
                                },
                                required = new[] { "list_elementIds" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_schedule_sorting_info",
                            description = "Возвращает информацию о правилах сортировки и группировки в спецификации Revit. " +
                                          "Выходные данные: scheduleId, scheduleName, hasSorting, sortLevelsCount, sorting, note. " +
                                          "Для КАЖДОГО УРОВНЯ СОРТИРОВКИ (sorting): " +
                                          "- level: уровень сортировки (1, 2, 3...) " +
                                          "- fieldId: ID поля сортировки " +
                                          "- fieldName: имя поля " +
                                          "- parameterId: ID параметра (если доступен) " +
                                          "- fieldType: тип поля " +
                                          "- sortOrder: направление сортировки (по возрастанию/убыванию) " +
                                          "- showBlankRow: показывать ли пустую строку " +
                                          "- showHeader: показывать ли заголовок группы " +
                                          "- showFooter: показывать ли итоги по группе " +
                                          "- showFooterCount: показывать ли количество в итогах группы " +
                                          "- showFooterTitle: показывать ли заголовок в итогах группы " +
                                          "- showGrandTotal: показывать ли общий итог " +
                                          "- showGrandTotalCount: показывать ли количество в общем итоге " +
                                          "- showGrandTotalTitle: показывать ли заголовок общего итога " +
                                          "- grandTitle: текст заголовка общего итога " +
                                          "- isItemized: надо ли группировать элементы попавшие в спецификацию (каждый элемент в отдельной строке) " +
                                          "Полезно для: анализа структуры спецификации, понимания порядка строк, проверки настроек группировки и итогов.",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    list_elementIds = new
                                    {
                                        type = "array",
                                        items = new { type = "integer" },
                                        description = "Список ID спецификаций (ViewSchedule)"
                                    }
                                },
                                required = new[] { "list_elementIds" }
                            }
                        }
                    },

                    //new
                    //{
                    //    type = "function",
                    //    function = new
                    //    {
                    //        name = "get_schedule_rows_with_elements",
                    //        description = "Возвращает все строки спецификации Revit с данными по каждому видимому столбцу и списком ID элементов, связанных со спецификацией. " +
                    //                      "Выходные данные: " +
                    //                      "- success: успешность выполнения " +
                    //                      "- schedule_id: ID спецификации " +
                    //                      "- schedule_name: имя спецификации " +
                    //                      "- is_itemized: детализирована ли спецификация (true = одна строка = один элемент) " +
                    //                      "- row_count: количество строк данных " +
                    //                      "- column_count: количество видимых столбцов " +
                    //                      "- total_columns: общее количество столбцов (включая скрытые) " +
                    //                      "- hidden_columns_count: количество скрытых столбцов " +
                    //                      "- headers: список заголовков видимых столбцов " +
                    //                      "- rows: список строк с данными. Для каждой строки: " +
                    //                      "    • row_index: индекс строки " +
                    //                      "    • values: словарь (заголовок столбца → значение ячейки) " +
                    //                      "- element_ids: список ID всех элементов Revit, связанных со спецификацией " +
                    //                      "Особенности: " +
                    //                      "- При запросе пользвотееля отсчет строк надо начинать от 1 (то есть 1 строка с данными начинается с 1 и далее 2, 3, 4), пропуская при этом заловок и пустые строки" +
                    //                      "- Скрытые столбцы автоматически исключаются из результата. " +
                    //                      "- Значения возвращаются в том виде, как их видит пользователь. " +
                    //                      "- Для детализированных спецификаций (is_itemized = true) порядок строк соответствует порядку element_ids. " +
                    //                      "- Для сгруппированных спецификаций (is_itemized = false) одна строка может соответствовать нескольким элементам. " +
                    //                      "Полезно для: экспорта данных спецификации, анализа содержимого, получения ID элементов для дальнейшей обработки.",
                    //        parameters = new
                    //        {
                    //            type = "object",
                    //            properties = new
                    //            {
                    //                scheduleId = new
                    //                {
                    //                    type = "integer",
                    //                    description = "ID спецификации (ViewSchedule)"
                    //                }
                    //            },
                    //            required = new[] { "scheduleId" }
                    //        }
                    //    }
                    //},



                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_if_elements_pass_filter",
                            description = "Проверяет, проходит ли каждый элемент из списка через условие заданного фильтра видов (ParameterFilterElement). " +
                                          "Фильтры видов используются в Revit для переопределения графики элементов на видах на основе их параметров. " +
                                          "Возвращает для каждого elementId булево значение: true — элемент соответствует правилам фильтра, false — не соответствует. " +
                                          "Полезно для: проверки корректности работы фильтра, поиска элементов, которые должны попадать под фильтр, " +
                                          "диагностики несоответствий в настройках фильтрации, аудита видимости элементов на видах. " +
                                          "Как это работает: метод использует ElementFilter.PassesFilter() для проверки каждого элемента. " +
                                          "Возвращает: filter_results (словарь elementId -> bool), count (общее количество), passed_count (количество прошедших), " +
                                          "failed_count (количество не прошедших), not_found_count (не найденные элементы), " +
                                          "filter_info (информация о фильтре: id фильтра, имя, список целевые категории и заданы ли правило в фильтре). " +
                                          "Важно: фильтр должен существовать в документе. ID фильтра можно получить через команды: " +
                                          "- get_graphic_filters_applied_to_views (получить фильтры, применённые к виду) " +
                                          "- get_all_parameter_filters (получить все фильтры в документе). " +
                                          "Пример использования: сначала получите ID фильтра из вида, затем проверьте конкретные элементы.",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    filterId = new
                                    {
                                        type = "integer",
                                        description = "Element id фильтра видов (ParameterFilterElement)"
                                    },
                                    list_elementIds = new
                                    {
                                        type = "array",
                                        items = new { type = "integer" },
                                        description = "Список element id элементов для проверки"
                                    }
                                },
                                required = new[] { "filterId", "list_elementIds" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "set_view_section_box_to_elements",
                            description = "Подрезает 3D вид по границам указанных элементов с отступом 500 мм. " +
                                          "Важно: перед вызовом команды пользователь должен открыть 3D вид. " +
                                          "Если открыт не 3D вид, команда вернёт ошибку с просьбой открыть 3D вид. " +
                                          "Команда автоматически включает Section Box на виде и устанавливает его границы " +
                                          "по расширенному bounding box'у переданных элементов. " +
                                          "Параметры: " +
                                          "- list_elementIds: список ID элементов, по которым будет подрезан вид " +
                                          "После успешного выполнения вид будет подрезан. " +
                                          "Возвращает: success (bool), message (строка с результатом), " +
                                          "elements_processed (количество обработанных элементов), " +
                                          "invalid_ids (список некорректных ID, если есть).",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    list_elementIds = new
                                    {
                                        type = "array",
                                        items = new { type = "integer" },
                                        description = "Список ID элементов, по которым нужно подрезать 3D вид"
                                    }
                                },
                                required = new[] { "list_elementIds" }
                            }
                        }
                    },

                    new
                    {
                        type = "function",
                        function = new
                        {
                            name = "get_journal_entries_since",
                            description = "Извлекает записи из журнала Revit начиная с указанной даты/времени. " +
                                          "Журналы Revit хранятся в %LOCALAPPDATA%\\Autodesk\\Revit\\<версия>\\Journals\\ " +
                                          "Команда автоматически находит самый свежий журнал и извлекает все записи от указанной даты до конца файла. " +
                                          "Входные параметры: " +
                                          "- dateTime: дата и время в формате (день.месяц.год час:минута). " +
                                          "  Поддерживаемые форматы: " +
                                          "  • '26 марта' или '26 марта 2026' (время = 00:00) " +
                                          "  • '26.03.2026' или '26.03.2026 14:30' " +
                                          "  • '26.03' (год = текущий, время = 00:00) " +
                                          "Если время не указано, используется время 00:00 сегодняшнего числа. " +
                                          "Возвращает: " +
                                          "- success: успешность выполнения " +
                                          "- journal_file: путь к файлу журнала " +
                                          "- journal_date: дата последнего изменения журнала " +
                                          "- target_date: обработанная целевая дата " +
                                          "- entries: текст журнала начиная с указанной даты " +
                                          "- entry_count: количество найденных записей " +
                                          "- total_size_bytes: размер извлечённого текста в байтах " +
                                          "- debug_info: ДИАГНОСТИЧЕСКАЯ ИНФОРМАЦИЯ (для отладки). Содержит: " +
                                          "    • Всего строк с датами в журнале " +
                                          "    • Первую дату в журнале " +
                                          "    • Последнюю дату в журнале " +
                                          "    • Целевую дату поиска " +
                                          "    • Был ли начат сбор записей " +
                                          "    • Количество найденных записей " +
                                          "  Если entry_count = 0, эта информация поможет понять причину: " +
                                          "  - целевая дата позже последней записи в журнале, " +
                                          "  - или в журнале нет записей с датами, " +
                                          "  - или формат даты в журнале не распознан. " +
                                          "- message: человеко-читаемое сообщение о результате " +
                                          "Важно: если извлечённый текст превышает 200 КБ, он автоматически обрезается. " +
                                          "Полезно для: анализа ошибок Revit, поиска проблем в сессии, аудита действий пользователя.",
                            parameters = new
                            {
                                type = "object",
                                properties = new
                                {
                                    dateTime = new
                                    {
                                        type = "string",
                                        description = "Дата и время начала извлечения. Примеры: '26 марта', '26.03.2026 14:30'"
                                    }
                                },
                                required = new[] { "dateTime" }
                            }
                        }
                    }



                    //new
                    //{
                    //    type = "function",
                    //    function = new
                    //    {
                    //        name = "get_document_switched",
                    //        description = "Переключает контекст всех последующих вызовов на связанный документ (Revit Link) или возвращает обратно к основному. " +
                    //                      "Позволяет исследовать содержимое вложенных файлов (связанных моделей Revit или IFC). " +
                    //                      "ВАЖНО: После переключения на связанный документ, все последующие команды будут работать с этим документом. " +
                    //                      "Для возврата к основному документу используйте switchMainDoc = true. " +
                    //                      "Параметры: " +
                    //                      "- elementId: ID элемента RevitLinkInstance (связи) для переключения на связанный документ " +
                    //                      "- switchMainDoc: true — вернуться в основной документ (игнорирует elementId) " +
                    //                      "Возвращает: " +
                    //                      "- success: успешность операции " +
                    //                      "- current_document: название текущего активного документа после переключения " +
                    //                      "- language_of_model: язык интерфейса переключённого документа " +
                    //                      "- link_info: информация о связи (имя файла, трансформация) " +
                    //                      "Примеры использования: " +
                    //                      "1. Переключиться на связанную модель: get_document_switched(elementId=12345) " +
                    //                      "2. Вернуться к основному документу: get_document_switched(switchMainDoc=true) " +
                    //                      "Полезно для: анализа элементов в связанных моделях, проверки коллизий, аудита ссылок, " +
                    //                      "работы с IFC-файлами, импортированными в Revit.",
                    //        parameters = new
                    //        {
                    //            type = "object",
                    //            properties = new
                    //            {
                    //                elementId = new
                    //                {
                    //                    type = "integer",
                    //                    description = "Element id RevitLinkInstance для переключения на связанный документ. По умолчанию -1."
                    //                },
                    //                switchMainDoc = new
                    //                {
                    //                    type = "boolean",
                    //                    description = "true — вернуться в основной документ. По умолчанию false."
                    //                }
                    //            },
                    //            required = new string[] { }
                    //        }
                    //    }
                    //}


















            //===========================================   конец   команд    ======================================        
        };



            var requestBody = new
            {
                model = _connectionType == ConnectionType.OnlineAPI ? "deepseek-v4-flash" : "qwen3-8b",
                messages = messagesWithSystem,
                temperature = 0.7,
                //max_tokens = 4000,
                tools = toolsArray
            };

            var json = Newtonsoft.Json.JsonConvert.SerializeObject(requestBody);
            var content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");

            var response = await _httpClient.PostAsync(apiUrl, content);
            var responseJson = await response.Content.ReadAsStringAsync();


            if (!response.IsSuccessStatusCode)
            {
                string errorMsg = $"HTTP {response.StatusCode}: {responseJson}";
                return errorMsg;  // ← Возвращаем ПОЛНУЮ ошибку, а не обрезанную!
            }

            if (!responseJson.TrimStart().StartsWith("{"))
            {
                return $"Не JSON: {responseJson.Substring(0, Math.Min(500, responseJson.Length))}...";
            }


            try
            {
                var jObject = JObject.Parse(responseJson); //проверка на способность парситься, если нет, то переход в блок catch

                try
                {
                    var usage = jObject["usage"];
                    if (usage != null)
                    {
                        // Получаем данные о токенах
                        int cacheHit = usage["prompt_cache_hit_tokens"]?.Value<int>() ?? 0;
                        int cacheMiss = usage["prompt_cache_miss_tokens"]?.Value<int>() ?? 0;
                        int completion = usage["completion_tokens"]?.Value<int>() ?? 0;
                        int total = usage["total_tokens"]?.Value<int>() ?? 0;

                        // Сохраняем в логгер (нужно будет передать вопрос и ответ)
                        // Пока сохраняем в поле класса, чтобы использовать позже
                        _lastCacheHit = cacheHit;
                        _lastCacheMiss = cacheMiss;
                        _lastCompletion = completion;
                        _lastTotal = total;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Ошибка логирования токенов: {ex.Message}");
                }

                return responseJson;
            }
            catch
            {
                return responseJson;  // Если не JSON — возвращаем как текст
            }
        }


        //Комманды для парсинга ответа============================================================================================================

        /// <summary>
        /// Парсит Markdown-разметку и добавляет форматированный текст в RichTextBox
        /// </summary>
        private void ParseMarkdownToRichTextBox(RichTextBox rtb, string markdown)
        {
            if (string.IsNullOrEmpty(markdown))
                return;

            rtb.Document = new FlowDocument();
            var lines = markdown.Split('\n');


            bool inTable = false;
            var tableRows = new List<string[]>();
            int tableColumnCount = 0;

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                // Проверяем, является ли строка частью таблицы (начинается с |)
                if (line.StartsWith("|") && line.EndsWith("|"))
                {
                    // Пропускаем строку-разделитель (|---------|------|)
                    if (System.Text.RegularExpressions.Regex.IsMatch(line, @"^\|\s*[\-:]+\s*\|"))
                    {
                        continue;
                    }

                    if (!inTable)
                    {
                        inTable = true;
                        tableRows.Clear();
                        tableColumnCount = 0;
                    }

                    // Разбиваем строку на ячейки
                    string[] cells = line.Trim('|').Split('|')
                        .Select(c => c.Trim())
                        .ToArray();

                    tableColumnCount = Math.Max(tableColumnCount, cells.Length);
                    tableRows.Add(cells);
                }
                else
                {
                    // Если вышли из таблицы, отображаем её
                    if (inTable && tableRows.Count > 0)
                    {
                        CreateTableInRichTextBox(rtb, tableRows, tableColumnCount);
                        inTable = false;
                        tableRows.Clear();
                    }

                    // Обрабатываем обычную строку
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        ProcessRegularLine(rtb, line);
                    }
                    else if (string.IsNullOrWhiteSpace(line))
                    {
                        rtb.Document.Blocks.Add(new Paragraph());
                    }
                }
            }

            // Если таблица была в конце документа
            if (inTable && tableRows.Count > 0)
            {
                CreateTableInRichTextBox(rtb, tableRows, tableColumnCount);
            }
        }

        /// <summary>
        /// Создаёт таблицу в RichTextBox с поддержкой форматирования внутри ячеек
        /// </summary>
        private void CreateTableInRichTextBox(RichTextBox rtb, List<string[]> rows, int columnCount)
        {
            var table = new Table();
            table.CellSpacing = 0;
            table.BorderBrush = Brushes.Gray;
            table.BorderThickness = new Thickness(1);
            table.Margin = new Thickness(0, 1, 0, 1);

            // Настраиваем ширину столбцов
            for (int i = 0; i < columnCount; i++)
            {
                table.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
            }

            for (int rowIdx = 0; rowIdx < rows.Count; rowIdx++)
            {
                var row = new TableRow();
                var cells = rows[rowIdx];

                // Выравнивание для заголовка (первая строка)
                if (rowIdx == 0)
                {
                    row.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(240, 240, 240));
                }

                for (int colIdx = 0; colIdx < columnCount; colIdx++)
                {
                    string cellText = colIdx < cells.Length ? cells[colIdx] : "";

                    // Создаём параграф с форматированием
                    var paragraph = new Paragraph();
                    ParseInlineMarkdown(paragraph, cellText);  // ← Поддерживает жирный, курсив, код

                    var cell = new TableCell(paragraph);
                    cell.BorderBrush = Brushes.Gray;
                    cell.BorderThickness = new Thickness(1);
                    cell.Padding = new Thickness(5);

                    // Для заголовка делаем текст жирным (дополнительно к форматированию)
                    if (rowIdx == 0)
                    {
                        cell.FontWeight = FontWeights.Bold;
                    }

                    row.Cells.Add(cell);
                }

                table.RowGroups.Add(new TableRowGroup());
                table.RowGroups.Last().Rows.Add(row);
            }

            rtb.Document.Blocks.Add(table);
        }

        /// <summary>
        /// Обрабатывает обычную строку текста (не таблицу)
        /// </summary>
        private void ProcessRegularLine(RichTextBox rtb, string line)
        {
            var para = new Paragraph();
            para.Margin = new Thickness(0, 0, 0, 5);

            if (line.StartsWith("### "))
            {
                var run = new Run(line.Substring(4));
                run.FontWeight = FontWeights.Bold;
                run.FontSize = 14;
                para.Inlines.Add(run);
                rtb.Document.Blocks.Add(para);
                return;
            }

            if (line.StartsWith("## "))
            {
                var run = new Run(line.Substring(3));
                run.FontWeight = FontWeights.Bold;
                run.FontSize = 16;
                para.Inlines.Add(run);
                rtb.Document.Blocks.Add(para);
                return;
            }

            if (line.StartsWith("# "))
            {
                var run = new Run(line.Substring(2));
                run.FontWeight = FontWeights.Bold;
                run.FontSize = 18;
                para.Inlines.Add(run);
                rtb.Document.Blocks.Add(para);
                return;
            }

            if (line.Trim() == "---" || line.Trim() == "***")
            {
                var separator = new Separator();
                para.Inlines.Add(new InlineUIContainer(separator));
                rtb.Document.Blocks.Add(para);
                return;
            }

            ParseInlineMarkdown(para, line);
            rtb.Document.Blocks.Add(para);
        }

        /// <summary>
        /// Обрабатывает Markdown-разметку внутри строки (жирный, курсив, код)
        /// </summary>
        private void ParseInlineMarkdown(Paragraph paragraph, string text)
        {
            int pos = 0;
            int length = text.Length;

            while (pos < length)
            {
                int boldStart = text.IndexOf("**", pos);
                int italicStart = text.IndexOf("*", pos);
                int codeStart = text.IndexOf("`", pos);

                int nextMarker = -1;
                string markerType = null;

                if (boldStart != -1 && (nextMarker == -1 || boldStart < nextMarker))
                {
                    nextMarker = boldStart;
                    markerType = "bold";
                }
                if (italicStart != -1 && (nextMarker == -1 || italicStart < nextMarker))
                {
                    if (italicStart + 1 < length && text[italicStart + 1] != '*')
                    {
                        nextMarker = italicStart;
                        markerType = "italic";
                    }
                }
                if (codeStart != -1 && (nextMarker == -1 || codeStart < nextMarker))
                {
                    nextMarker = codeStart;
                    markerType = "code";
                }

                if (nextMarker == -1)
                {
                    if (pos < length)
                    {
                        paragraph.Inlines.Add(new Run(text.Substring(pos)));
                    }
                    break;
                }

   
                if (nextMarker > pos)
                {
                    paragraph.Inlines.Add(new Run(text.Substring(pos, nextMarker - pos)));
                }

                pos = nextMarker;

                switch (markerType)
                {
                    case "bold":
                        int boldEnd = text.IndexOf("**", pos + 2);
                        if (boldEnd != -1)
                        {
                            var run = new Run(text.Substring(pos + 2, boldEnd - pos - 2));
                            run.FontWeight = FontWeights.Bold;
                            paragraph.Inlines.Add(run);
                            pos = boldEnd + 2;
                        }
                        else
                        {
                            paragraph.Inlines.Add(new Run(text.Substring(pos)));
                            pos = length;
                        }
                        break;

                    case "italic":
                        int italicEnd = text.IndexOf("*", pos + 1);
                        if (italicEnd != -1)
                        {
                            var run = new Run(text.Substring(pos + 1, italicEnd - pos - 1));
                            run.FontStyle = FontStyles.Italic;
                            paragraph.Inlines.Add(run);
                            pos = italicEnd + 1;
                        }
                        else
                        {
                            paragraph.Inlines.Add(new Run(text.Substring(pos)));
                            pos = length;
                        }
                        break;

                    case "code":
                        int codeEnd = text.IndexOf("`", pos + 1);
                        if (codeEnd != -1)
                        {
                            var run = new Run(text.Substring(pos + 1, codeEnd - pos - 1));
                            run.FontFamily = new System.Windows.Media.FontFamily("Consolas");
                            run.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 112, 192)); 
                            paragraph.Inlines.Add(run);
                            pos = codeEnd + 1;
                        }
                        break;
                }
            }
        }


    }
}
