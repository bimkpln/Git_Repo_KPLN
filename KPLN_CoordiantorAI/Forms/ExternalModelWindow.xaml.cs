using System;
using System.Collections.Generic;
using System.Diagnostics;
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
using KPLN_CoordiantorAI.ExternalAIModel.Mcp;
using KPLN_CoordiantorAI.ExternalAIModel.McpClient;
using KPLN_CoordiantorAI.ExternalModel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Windows.Controls.Primitives;
using static KPLN_CoordiantorAI.ExternalModel.Commands;
using WpfGrid = System.Windows.Controls.Grid;
using WpfTextBox = System.Windows.Controls.TextBox;
using Control = System.Windows.Controls.Control;
using KPLN_CoordiantorAI.ExternalAIModel;
using System.Threading;

namespace KPLN_CoordiantorAI.Forms
{
    public interface IModelProgressReporter
    {
        void Report(string status);
        void Clear();
    }

    internal sealed class WpfModelProgressReporter : IModelProgressReporter
    {
        private readonly ExternalModelControl _owner;

        public WpfModelProgressReporter(ExternalModelControl owner)
        {
            _owner = owner;
        }

        public void Report(string status)
        {
            if (_owner != null)
                _owner.ReportProgressStatus(status);
        }

        public void Clear()
        {
            if (_owner != null)
                _owner.ClearProgressStatus();
        }
    }

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
        private ChatLogger _logger;
        private DiagnosticLogger _diagnosticLogger;
        private ConnectionType _connectionType;
        private ExternalModelSettings _settings;
        private Window _hostWindow;
        private string _currentDiagnosticRequestId;
        private Stopwatch _currentRequestStopwatch;
        private bool _isRequestInProgress;
        private bool _isClosing;
        private CancellationTokenSource _currentRequestCancellation;
        private const int MaxRealToolCallsPerRound = 5;
        private const string DefaultProgressStatus = "Анализирую запрос";
        private IModelProgressReporter _progressReporter;






        private int _lastCacheHit = 0;
        private int _lastCacheMiss = 0;
        private int _lastCompletion = 0;
        private int _lastTotal = 0;


        private readonly HttpClient _httpClient = new HttpClient();
        private readonly Dictionary<string, int> _currentToolAreaStats = new Dictionary<string, int>(StringComparer.Ordinal);
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
            _connectionType = connectionType;
            _settings = settings ?? new ExternalModelSettings();
            _progressReporter = new WpfModelProgressReporter(this);


            // Инициализация логгера
            _logger = new ChatLogger(_settings.LogFolder, "wpf_window");
            //Можно задать путь через настройки к логу-диганостики
            _diagnosticLogger = new DiagnosticLogger(null, "wpf_window");

            // Анимация точек
            _typingTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(400) };
            _typingTimer.Tick += TypingTimer_Tick;

            Loaded += OnExternalModelControlLoaded;
            Unloaded += OnExternalModelControlUnloaded;
            SetupEventHandlers();

            // Показываем пользователю, какой режим активен
            DisplayTitle = GetModelTitle();
            TitleTextBlock.Text = DisplayTitle;
        }

        // Показываем пользователю, какой режим активен
        private string GetModelTitle()
        {
            string mode = _connectionType == ConnectionType.OnlineAPI
                ? "Online (API key)"
                : "Local (LM Studio)";

            return $"Работа с моделью - {_doc.Title} - {mode}";
        }

        private void OnExternalModelControlLoaded(object sender, RoutedEventArgs e)
        {
            _isClosing = false;
            _hostWindow = Window.GetWindow(this);
            if (_hostWindow == null)
                return;

            _hostWindow.Closing -= OnHostWindowClosing;
            _hostWindow.Closed -= OnHostWindowClosed;
            _hostWindow.Closing += OnHostWindowClosing;
            _hostWindow.Closed += OnHostWindowClosed;

            _diagnosticLogger.LogEvent(null, "WINDOW.LOADED", new Dictionary<string, object>
            {
                { "title", _hostWindow.Title },
                { "model", GetCurrentRevitModelName() },
                { "view", GetCurrentRevitViewName() }
            });
        }

        private void OnExternalModelControlUnloaded(object sender, RoutedEventArgs e)
        {
            if (_hostWindow == null)
                return;

            _hostWindow.Closing -= OnHostWindowClosing;
            _hostWindow.Closed -= OnHostWindowClosed;
            _hostWindow = null;
        }

        private void OnHostWindowClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            _isClosing = true;
            _diagnosticLogger.LogEvent(_currentDiagnosticRequestId, "WINDOW.CLOSING", new Dictionary<string, object>
            {
                { "isRequestInProgress", _isRequestInProgress },
                { "elapsedMs", GetCurrentRequestElapsedMs() }
            });
            CancelCurrentRequest("windowClosing");
        }

        private void OnHostWindowClosed(object sender, EventArgs e)
        {
            _isClosing = true;
            _diagnosticLogger.LogEvent(_currentDiagnosticRequestId, "WINDOW.CLOSED", new Dictionary<string, object>
            {
                { "isRequestInProgress", _isRequestInProgress },
                { "elapsedMs", GetCurrentRequestElapsedMs() }
            });
            CancelCurrentRequest("windowClosed");
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
            _currentTypingFrame = (_currentTypingFrame + 1) % 4;
            UpdateProgressIndicatorText();
        }

        private string _currentProgressStatus = DefaultProgressStatus;

        private string GetAnimatedProgressText()
        {
            string dots = new string('.', _currentTypingFrame);
            return (_currentProgressStatus ?? DefaultProgressStatus) + dots;
        }

        private void UpdateProgressIndicatorText()
        {
            if (_typingIndicator?.Child is RichTextBox richTextBox)
            {
                var paragraph = new Paragraph();
                paragraph.Inlines.Add(new Run(GetAnimatedProgressText()));
                richTextBox.Document = new FlowDocument(paragraph);
            }
        }

        //показать загрузку ИИ
        private void ShowTypingIndicator(string status = null)
        {
            if (_typingIndicator != null)
                ChatHistory.Children.Remove(_typingIndicator);

            _currentProgressStatus = string.IsNullOrWhiteSpace(status) ? DefaultProgressStatus : status;
            _currentTypingFrame = 0;
            _typingIndicator = CreateTypingIndicator();
            ChatHistory.Children.Add(_typingIndicator);
            ChatScrollViewer.ScrollToEnd();

            UpdateProgressIndicatorText();
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

        internal void ReportProgressStatus(string status)
        {
            if (string.IsNullOrWhiteSpace(status))
                return;

            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.BeginInvoke(new Action(() => ReportProgressStatus(status)));
                return;
            }

            _currentProgressStatus = status.Trim().TrimEnd('.');
            _currentTypingFrame = 0;

            if (_typingIndicator == null)
            {
                ShowTypingIndicator(_currentProgressStatus);
                return;
            }

            UpdateProgressIndicatorText();
            ChatScrollViewer.ScrollToEnd();
        }

        internal void ClearProgressStatus()
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.BeginInvoke(new Action(ClearProgressStatus));
                return;
            }

            HideTypingIndicator();
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

        private static string ExtractAiErrorMessage(string responseJson)
        {
            if (string.IsNullOrWhiteSpace(responseJson))
                return "AI API returned an empty error response.";

            try
            {
                JObject obj = JObject.Parse(responseJson);
                JToken error = obj["error"];
                if (error == null)
                    return responseJson;

                JToken message = error["message"];
                if (message != null && !string.IsNullOrWhiteSpace(message.ToString()))
                    return message.ToString();

                return JsonConvert.SerializeObject(error);
            }
            catch
            {
                return TruncateForDiagnostics(responseJson, 1000);
            }
        }

        //процесс отправки сообщения
        public async Task SendMessage(string revitContext = "")
        {
            string userMessage = InputTextBox.Text.Trim();
            if (string.IsNullOrEmpty(userMessage)) return;
            if (_isClosing) return;

            DateTime requestTime = DateTime.Now;
            string requestModelName = GetCurrentRevitModelName();
            string requestViewName = GetCurrentRevitViewName();
            _currentToolAreaStats.Clear();
            string requestId = Guid.NewGuid().ToString("N");
            _currentDiagnosticRequestId = requestId;
            _currentRequestStopwatch = Stopwatch.StartNew();
            _isRequestInProgress = true;
            _currentRequestCancellation?.Dispose();
            _currentRequestCancellation = new CancellationTokenSource();
            CancellationTokenSource requestCancellation = _currentRequestCancellation;
            CancellationToken cancellationToken = requestCancellation.Token;
            _diagnosticLogger.LogEvent(requestId, "SendMessage.START", new Dictionary<string, object>
            {
                { "textLength", userMessage.Length },
                { "textPreview", TrimForDiagnostics(userMessage, 300) },
                { "model", requestModelName },
                { "view", requestViewName },
                { "connectionType", _connectionType },
                { "messagesBeforeAdd", ChatHistoryMessages.Count }
            });

            // Пользовательское сообщение
            var userMsg = new { role = "user", content = userMessage };
            ChatHistoryMessages.Add(userMsg);

            ChatHistory.Children.Add(CreateMessageBlock($"Вы: {userMessage}", true));
            InputTextBox.Clear();
            SendButton.IsEnabled = false;

            try
            {
                _progressReporter.Report("Анализирую запрос");

                _diagnosticLogger.LogEvent(requestId, "AI_REQUEST.START", new Dictionary<string, object>
                {
                    { "phase", "initial" },
                    { "messages", ChatHistoryMessages.Count },
                    { "toolsEnabled", true }
                });
                _progressReporter.Report("Проверяю, нужны ли данные из Revit");
                string response = await SendToOpenRouter(ChatHistoryMessages, cancellationToken);
                cancellationToken.ThrowIfCancellationRequested();
                if (IsCloseOrCancellationRequested(cancellationToken)) return;
                _progressReporter.Report("Обрабатываю результат");
                _diagnosticLogger.LogEvent(requestId, "AI_REQUEST.END", new Dictionary<string, object>
                {
                    { "phase", "initial" },
                    { "responseLength", response == null ? 0 : response.Length }
                });

                if (!response.TrimStart().StartsWith("{"))
                {
                    _diagnosticLogger.LogEvent(requestId, "AI_RESPONSE.NON_JSON", new Dictionary<string, object>
                    {
                        { "responsePreview", TrimForDiagnostics(response, 500) }
                    });
                    ChatHistory.Children.Add(CreateMessageBlock($"❌ AI: {response}", false));


                    if (ChatHistoryMessages.Count > 0 && ChatHistoryMessages.Last().GetType().GetProperty("role")?.GetValue(ChatHistoryMessages.Last())?.ToString() == "tool")
                    {
                        ChatHistoryMessages.RemoveAt(ChatHistoryMessages.Count - 1);
                    }


                    return;
                }

                // Парсим ответ ИИ
                var responseJObject = JObject.Parse(response);
                _diagnosticLogger.LogEvent(requestId, "AI_RESPONSE.PARSED", GetResponseDiagnostics(responseJObject));

                while (true)
                {
                    var message = responseJObject["choices"]?[0]?["message"].ToObject<JObject>();
                    var toolCalls = message?["tool_calls"] as JArray;


                    ChatHistoryMessages.Add(message.ToObject<object>());


                    // Если есть tool_calls — выполняем
                    if (toolCalls != null && toolCalls.Count > 0)
                    {
                        Stopwatch toolsBatchStopwatch = Stopwatch.StartNew();
                        _diagnosticLogger.LogEvent(requestId, "TOOLS_BATCH.START", new Dictionary<string, object>
                        {
                            { "toolCalls", toolCalls.Count },
                            { "maxRealToolCallsPerRound", MaxRealToolCallsPerRound }
                        });
                        int realToolCallsExecuted = 0;
                        int toolCallsSkipped = 0;
                        foreach (JObject tc in toolCalls)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            if (IsCloseOrCancellationRequested(cancellationToken)) return;

                            if (realToolCallsExecuted < MaxRealToolCallsPerRound)
                            {
                                await ProcessSingleToolCall(tc);
                                realToolCallsExecuted++;
                            }
                            else
                            {
                                AddSkippedToolCallResultToHistory(tc, toolCalls.Count, realToolCallsExecuted);
                                toolCallsSkipped++;
                            }
                        }
                        toolsBatchStopwatch.Stop();
                        _diagnosticLogger.LogEvent(requestId, "TOOLS_BATCH.END", new Dictionary<string, object>
                        {
                            { "toolCalls", toolCalls.Count },
                            { "realToolCallsExecuted", realToolCallsExecuted },
                            { "toolCallsSkipped", toolCallsSkipped },
                            { "elapsedMs", toolsBatchStopwatch.ElapsedMilliseconds }
                        });

                        // После всех tools → следующий запрос к ИИ
                        _progressReporter.Report("Формирую ответ");
                        _diagnosticLogger.LogEvent(requestId, "AI_REQUEST.START", new Dictionary<string, object>
                        {
                            { "phase", "afterTools" },
                            { "messages", ChatHistoryMessages.Count },
                            { "toolsEnabled", true }
                        });
                        response = await SendToOpenRouter(ChatHistoryMessages, cancellationToken);
                        cancellationToken.ThrowIfCancellationRequested();
                        if (IsCloseOrCancellationRequested(cancellationToken)) return;
                        _progressReporter.Report("Обрабатываю результат");
                        _diagnosticLogger.LogEvent(requestId, "AI_REQUEST.END", new Dictionary<string, object>
                        {
                            { "phase", "afterTools" },
                            { "responseLength", response == null ? 0 : response.Length }
                        });
                        responseJObject = JObject.Parse(response);
                        _diagnosticLogger.LogEvent(requestId, "AI_RESPONSE.PARSED", GetResponseDiagnostics(responseJObject));

                        if (!IsSuccessResponse(response))
                        {
                            // Если ошибка - показываем и выходим
                            ChatHistory.Children.Add(CreateMessageBlock($"Ошибка после tools: {ExtractAiErrorMessage(response)}", false));
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
                            _progressReporter.Report("Формирую ответ");
                            DateTime responseTime = DateTime.Now;
                            _diagnosticLogger.LogEvent(requestId, "FINAL_RESPONSE.READY", new Dictionary<string, object>
                            {
                                { "contentLength", finalContent.Length },
                                { "elapsedMs", GetCurrentRequestElapsedMs() },
                                { "toolAreaStats", FormatToolAreaStatsForDiagnostics() },
                                { "cacheHitTokens", _lastCacheHit },
                                { "cacheMissTokens", _lastCacheMiss },
                                { "completionTokens", _lastCompletion },
                                { "totalTokens", _lastTotal }
                            });
                            _diagnosticLogger.LogEvent(requestId, "UI_RENDER.START", new Dictionary<string, object>
                            {
                                { "contentLength", finalContent.Length }
                            });
                            ChatHistory.Children.Add(CreateMessageBlock($"AI: {finalContent}", false));
                            _diagnosticLogger.LogEvent(requestId, "UI_RENDER.END");


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

                            _diagnosticLogger.LogEvent(requestId, "CHAT_LOG.START");
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
                                requestViewName,
                                GetCurrentToolAreaStatsForLog());
                            _diagnosticLogger.LogEvent(requestId, "CHAT_LOG.END");
                            break;
                        }
                    }
                }

            }
            catch (OperationCanceledException ex)
            {
                _progressReporter.Report("Запрос отменен");
                _diagnosticLogger.LogException(requestId, "SendMessage.CANCELED", ex, new Dictionary<string, object>
                {
                    { "elapsedMs", GetCurrentRequestElapsedMs() },
                    { "isClosing", _isClosing }
                });

                if (!IsCloseOrCancellationRequested(cancellationToken))
                    ChatHistory.Children.Add(CreateMessageBlock("Запрос отменен.", false));
            }
            catch (Exception ex)
            {
                _progressReporter.Report("Не удалось обработать запрос");
                _diagnosticLogger.LogException(requestId, "SendMessage.ERROR", ex, new Dictionary<string, object>
                {
                    { "elapsedMs", GetCurrentRequestElapsedMs() }
                });
                if (!IsCloseOrCancellationRequested(cancellationToken))
                    ChatHistory.Children.Add(CreateMessageBlock($"Ошибка: {ex.Message}", false));
                // Логируем ошибку (вопрос пользователя и текст ошибки)
                _logger.Log(userMessage, $"ОШИБКА: {ex.Message}", requestTime, DateTime.Now, requestModelName, requestViewName, GetCurrentToolAreaStatsForLog());
            }
            finally
            {
                _diagnosticLogger.LogEvent(requestId, "SendMessage.FINALLY", new Dictionary<string, object>
                {
                    { "elapsedMs", GetCurrentRequestElapsedMs() },
                    { "toolAreaStats", FormatToolAreaStatsForDiagnostics() }
                });
                _isRequestInProgress = false;
                if (_currentRequestStopwatch != null)
                    _currentRequestStopwatch.Stop();
                if (ReferenceEquals(_currentRequestCancellation, requestCancellation))
                {
                    _currentRequestCancellation.Dispose();
                    _currentRequestCancellation = null;
                }

                if (!IsCloseOrCancellationRequested(cancellationToken))
                {
                    _progressReporter.Clear();
                    SendButton.IsEnabled = true;
                    ChatScrollViewer.ScrollToEnd();
                }
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

        private long GetCurrentRequestElapsedMs()
        {
            return _currentRequestStopwatch == null ? 0 : _currentRequestStopwatch.ElapsedMilliseconds;
        }

        private bool IsCloseOrCancellationRequested(CancellationToken cancellationToken)
        {
            return _isClosing || cancellationToken.IsCancellationRequested || !IsLoaded;
        }

        private void CancelCurrentRequest(string reason)
        {
            CancellationTokenSource cancellation = _currentRequestCancellation;
            if (cancellation == null || cancellation.IsCancellationRequested)
                return;

            _diagnosticLogger.LogEvent(_currentDiagnosticRequestId, "REQUEST.CANCEL", new Dictionary<string, object>
            {
                { "reason", reason },
                { "isRequestInProgress", _isRequestInProgress },
                { "elapsedMs", GetCurrentRequestElapsedMs() }
            });

            try
            {
                cancellation.Cancel();
            }
            catch (ObjectDisposedException)
            {
            }
        }

        private static string TrimForDiagnostics(string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;

            string normalized = value.Replace("\r", " ").Replace("\n", " ").Trim();
            if (normalized.Length <= maxLength)
                return normalized;

            return normalized.Substring(0, maxLength) + "...";
        }

        private static Dictionary<string, object> GetResponseDiagnostics(JObject responseJObject)
        {
            JObject message = responseJObject?["choices"]?[0]?["message"] as JObject;
            JArray toolCalls = message?["tool_calls"] as JArray;
            string content = message?["content"]?.ToString() ?? string.Empty;

            return new Dictionary<string, object>
            {
                { "hasChoices", responseJObject?["choices"] != null },
                { "hasMessage", message != null },
                { "toolCalls", toolCalls == null ? 0 : toolCalls.Count },
                { "contentLength", content.Length }
            };
        }

        private string FormatToolAreaStatsForDiagnostics()
        {
            Dictionary<string, int> stats = GetCurrentToolAreaStatsForLog();
            if (stats.Count == 0)
                return "";

            return string.Join(",", stats.Select(i => i.Key + ":" + i.Value));
        }

        private void RegisterToolAreaCall(string toolName)
        {
            string area = GetToolAreaName(toolName);

            int currentCount;
            _currentToolAreaStats.TryGetValue(area, out currentCount);
            _currentToolAreaStats[area] = currentCount + 1;
        }

        private static string GetToolAreaName(string toolName)
        {
            RevitMcpToolDefinition definition;
            if (RevitMcpToolRegistry.TryGet(toolName, out definition) && !string.IsNullOrWhiteSpace(definition.Area))
                return definition.Area;

            return "Other";
        }

        private static string GetToolProgressStatus(string toolName)
        {
            switch (toolName)
            {
                case "get_active_view_in_revit":
                    return "Проверяю активный вид";
                case "get_model_categories":
                    return "Получаю категории модели";
                case "get_elements_by_category":
                    return "Получаю элементы выбранной категории";
                case "get_all_elements_shown_in_view":
                    return "Получаю элементы активного вида";
                case "get_journal_entries_since":
                    return "Читаю журнал Revit";
                case "set_user_selection_in_revit":
                    return "Выделяю элементы в Revit";
                case "set_view_section_box_to_elements":
                    return "Настраиваю 3D-вид";
                default:
                    return "Выполняю команду Revit";
            }
        }

        private Dictionary<string, int> GetCurrentToolAreaStatsForLog()
        {
            Dictionary<string, int> stats = new Dictionary<string, int>(StringComparer.Ordinal);
            IEnumerable<string> areas = RevitMcpToolRegistry.GetAll()
                .Select(tool => tool.Area)
                .Where(area => !string.IsNullOrWhiteSpace(area))
                .Concat(new[] { "Other" })
                .Distinct(StringComparer.Ordinal);
            foreach (string area in areas)
            {
                int count;
                if (_currentToolAreaStats.TryGetValue(area, out count) && count > 0)
                    stats[area] = count;
            }

            return stats;
        }





        private async Task ProcessSingleToolCall(JObject toolCall)
        {
            string toolName = toolCall["function"]?["name"]?.ToString();
            RegisterToolAreaCall(toolName);
            _progressReporter.Report(GetToolProgressStatus(toolName));
            string toolCallId = toolCall["id"]?.ToString() ?? Guid.NewGuid().ToString();

            string toolResult = "";
            var argsJson = toolCall["function"]?["arguments"]?.ToString() ?? "{}";
            var argsObj = JObject.Parse(argsJson);  // ← ТУТ объявляем!
            Stopwatch toolStopwatch = Stopwatch.StartNew();
            _diagnosticLogger.LogEvent(_currentDiagnosticRequestId, "TOOL.START", new Dictionary<string, object>
            {
                { "toolName", toolName },
                { "argsLength", argsJson.Length },
                { "toolArea", GetToolAreaName(toolName) }
            });

            try
            {
                toolResult = await ExecuteRegisteredMcpToolAsync(toolName, argsObj);
                AddToolResultToHistory(toolCallId, toolResult);
                toolStopwatch.Stop();
                _diagnosticLogger.LogEvent(_currentDiagnosticRequestId, "TOOL.END", new Dictionary<string, object>
                {
                    { "toolName", toolName },
                    { "resultLength", toolResult == null ? 0 : toolResult.Length },
                    { "elapsedMs", toolStopwatch.ElapsedMilliseconds },
                    { "executionPath", "mcp" }
                });
                _progressReporter.Report("Обрабатываю результат команды Revit");
            }
            catch (Exception ex)
            {
                _progressReporter.Report("Не удалось выполнить команду Revit");
                toolStopwatch.Stop();
                _diagnosticLogger.LogException(_currentDiagnosticRequestId, "TOOL.ERROR", ex, new Dictionary<string, object>
                {
                    { "toolName", toolName },
                    { "elapsedMs", toolStopwatch.ElapsedMilliseconds }
                });

                toolResult = JsonConvert.SerializeObject(new
                {
                    success = false,
                    error = ex.Message,
                    exception_type = ex.GetType().FullName,
                    tool_name = toolName
                });

                ChatHistoryMessages.Add(new
                {
                    role = "tool",
                    tool_call_id = toolCallId,
                    content = toolResult
                });

                _diagnosticLogger.LogEvent(_currentDiagnosticRequestId, "TOOL.ERROR_RESULT_ADDED", new Dictionary<string, object>
                {
                    { "toolName", toolName },
                    { "toolCallId", toolCallId },
                    { "resultLength", toolResult == null ? 0 : toolResult.Length }
                });
            }
        }

        private async Task<string> ExecuteRegisteredMcpToolAsync(string toolName, JObject arguments)
        {
            RevitMcpDiagnosticLogger.Log("WPF chat routes tool through MCP. ToolName=" + (toolName ?? "<null>"));

            using (InternalMcpClient mcpClient = new InternalMcpClient())
            {
                JObject mcpResult = await mcpClient.CallToolAsync(
                    toolName,
                    arguments ?? new JObject(),
                    CancellationToken.None);

                JToken structuredContent = mcpResult["structuredContent"];
                if (structuredContent != null)
                    return JsonConvert.SerializeObject(structuredContent);

                JArray content = mcpResult["content"] as JArray;
                if (content != null && content.Count > 0)
                {
                    JObject firstContent = content[0] as JObject;
                    if (firstContent != null && firstContent["text"] != null)
                        return firstContent["text"].ToString();
                }

                return JsonConvert.SerializeObject(mcpResult);
            }
        }

        private async Task<JArray> GetMcpOpenAiCompatibleToolsAsync(CancellationToken cancellationToken)
        {
            using (InternalMcpClient mcpClient = new InternalMcpClient())
            {
                InternalMcpToolProvider toolProvider = new InternalMcpToolProvider(mcpClient);
                JArray tools = await toolProvider.GetOpenAiCompatibleToolsAsync(cancellationToken);
                return tools;
            }
        }

        private void AddToolResultToHistory(string toolCallId, string toolResult)
        {
            ChatHistoryMessages.Add(new
            {
                role = "tool",
                tool_call_id = toolCallId,
                content = toolResult
            });
        }

        private void AddSkippedToolCallResultToHistory(JObject toolCall, int totalToolCallsInRound, int realToolCallsExecuted)
        {
            string toolName = toolCall["function"]?["name"]?.ToString();
            string toolCallId = toolCall["id"]?.ToString() ?? Guid.NewGuid().ToString();
            string toolResult = JsonConvert.SerializeObject(new
            {
                success = false,
                skipped = true,
                error = "Per-round tool execution limit reached.",
                message = "Only the first 5 tool calls from one model response are executed. Use the results already returned and request the next batch of tool calls separately if more data is needed.",
                tool_name = toolName,
                total_tool_calls_in_round = totalToolCallsInRound,
                real_tool_calls_executed = realToolCallsExecuted,
                max_real_tool_calls_per_round = MaxRealToolCallsPerRound
            });

            AddToolResultToHistory(toolCallId, toolResult);

            _diagnosticLogger.LogEvent(_currentDiagnosticRequestId, "TOOL.SKIPPED_PER_ROUND_LIMIT", new Dictionary<string, object>
            {
                { "toolName", toolName },
                { "toolCallId", toolCallId },
                { "totalToolCallsInRound", totalToolCallsInRound },
                { "realToolCallsExecuted", realToolCallsExecuted },
                { "maxRealToolCallsPerRound", MaxRealToolCallsPerRound },
                { "resultLength", toolResult.Length }
            });
        }



        private Border CreateMessageBlock(string text, bool isUser)
        {
            _diagnosticLogger.LogEvent(_currentDiagnosticRequestId, "CreateMessageBlock.START", new Dictionary<string, object>
            {
                { "isUser", isUser },
                { "textLength", text == null ? 0 : text.Length }
            });
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


                _diagnosticLogger.LogEvent(_currentDiagnosticRequestId, "MarkdownParse.START", new Dictionary<string, object>
                {
                    { "textLength", text == null ? 0 : text.Length }
                });
                // Парсим Markdown и заполняем RichTextBox
                ParseMarkdownToRichTextBox(richTextBox, text);

                _diagnosticLogger.LogEvent(_currentDiagnosticRequestId, "MarkdownParse.END");
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
            paragraph.Inlines.Add(new Run(GetAnimatedProgressText()));
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


        private async Task<string> SendToOpenRouter(List<object> _messages, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
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

            systemPrompt += "\r\nЕсли часть tool-вызовов вернулась с skipped=true и ошибкой Per-round tool execution limit reached, значит за один ответ модели было запрошено слишком много команд. Используй уже полученные результаты и запроси следующую небольшую пачку tool-вызовов отдельным шагом, если данных недостаточно.";

            var messagesWithSystem = new List<object> { new { role = "system", content = systemPrompt } };
            messagesWithSystem.AddRange(_messages);

            // Tool definitions are loaded from the local MCP tools/list endpoint.




            JArray mcpToolsArray = await GetMcpOpenAiCompatibleToolsAsync(cancellationToken);

            string modelName = _connectionType == ConnectionType.OnlineAPI ? "deepseek-v4-flash" : "qwen3-8b";
            var requestBody = new
            {
                model = modelName,
                messages = messagesWithSystem,
                temperature = 0.7,
                //max_tokens = 4000,
                tools = mcpToolsArray
            };

            var json = Newtonsoft.Json.JsonConvert.SerializeObject(requestBody);
            var content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");

            Stopwatch httpStopwatch = Stopwatch.StartNew();
            _diagnosticLogger.LogEvent(_currentDiagnosticRequestId, "HTTP.START", new Dictionary<string, object>
            {
                { "apiUrl", apiUrl },
                { "model", modelName },
                { "connectionType", _connectionType },
                { "jsonLength", json.Length },
                { "messages", messagesWithSystem.Count },
                { "tools", mcpToolsArray.Count },
                { "toolsSource", "mcp_tools_list" },
                { "timeoutSeconds", _httpClient.Timeout.TotalSeconds }
            });

            HttpResponseMessage response = null;
            string responseJson = string.Empty;
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                response = await _httpClient.PostAsync(apiUrl, content, cancellationToken);
                cancellationToken.ThrowIfCancellationRequested();
                responseJson = await response.Content.ReadAsStringAsync();
                cancellationToken.ThrowIfCancellationRequested();
                httpStopwatch.Stop();
                _diagnosticLogger.LogEvent(_currentDiagnosticRequestId, "HTTP.END", new Dictionary<string, object>
                {
                    { "statusCode", (int)response.StatusCode },
                    { "isSuccess", response.IsSuccessStatusCode },
                    { "responseLength", responseJson == null ? 0 : responseJson.Length },
                    { "elapsedMs", httpStopwatch.ElapsedMilliseconds }
                });
            }
            catch (OperationCanceledException ex)
            {
                httpStopwatch.Stop();
                _diagnosticLogger.LogException(_currentDiagnosticRequestId, "HTTP.CANCELED", ex, new Dictionary<string, object>
                {
                    { "elapsedMs", httpStopwatch.ElapsedMilliseconds }
                });
                throw;
            }
            catch (Exception ex)
            {
                httpStopwatch.Stop();
                _diagnosticLogger.LogException(_currentDiagnosticRequestId, "HTTP.ERROR", ex, new Dictionary<string, object>
                {
                    { "elapsedMs", httpStopwatch.ElapsedMilliseconds }
                });
                throw;
            }


            if (!response.IsSuccessStatusCode)
            {
                string message = BuildHttpApiErrorMessage(response, responseJson, json.Length);
                return JsonConvert.SerializeObject(new
                {
                    error = new
                    {
                        code = "ai_http_error",
                        message = message,
                        statusCode = (int)response.StatusCode,
                        requestJsonLength = json.Length,
                        responsePreview = TruncateForDiagnostics(responseJson, 1000)
                    }
                });
            }

            if (!responseJson.TrimStart().StartsWith("{"))
            {
                return JsonConvert.SerializeObject(new
                {
                    error = new
                    {
                        code = "ai_non_json_response",
                        message = "AI API returned a non-JSON response.",
                        requestJsonLength = json.Length,
                        responsePreview = TruncateForDiagnostics(responseJson, 1000)
                    }
                });
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
                        _lastCacheHit += cacheHit;
                        _lastCacheMiss += cacheMiss;
                        _lastCompletion += completion;
                        _lastTotal += total;
                        _diagnosticLogger.LogEvent(_currentDiagnosticRequestId, "USAGE.READ", new Dictionary<string, object>
                        {
                            { "cacheHit", cacheHit },
                            { "cacheMiss", cacheMiss },
                            { "completion", completion },
                            { "total", total },
                            { "accumulatedCacheHit", _lastCacheHit },
                            { "accumulatedCacheMiss", _lastCacheMiss },
                            { "accumulatedCompletion", _lastCompletion },
                            { "accumulatedTotal", _lastTotal }
                        });
                    }
                    else
                    {
                        _diagnosticLogger.LogEvent(_currentDiagnosticRequestId, "USAGE.MISSING");
                    }
                }
                catch (Exception ex)
                {
                    _diagnosticLogger.LogException(_currentDiagnosticRequestId, "USAGE.ERROR", ex);
                    System.Diagnostics.Debug.WriteLine($"Ошибка логирования токенов: {ex.Message}");
                }

                return responseJson;
            }
            catch (Exception ex)
            {
                _diagnosticLogger.LogException(_currentDiagnosticRequestId, "JSON_PARSE.ERROR", ex, new Dictionary<string, object>
                {
                    { "responsePreview", TrimForDiagnostics(responseJson, 500) }
                });
                return responseJson;  // Если не JSON — возвращаем как текст
            }
        }

        private static string BuildHttpApiErrorMessage(HttpResponseMessage response, string responseJson, int requestJsonLength)
        {
            string statusText = response == null
                ? "unknown"
                : ((int)response.StatusCode).ToString() + " " + response.StatusCode;

            string message = "AI API returned HTTP " + statusText + ".";
            if (requestJsonLength > 1000000)
                message += " The request payload is very large and may exceed the model/provider context or request-size limit.";

            string preview = TruncateForDiagnostics(responseJson, 500);
            if (!string.IsNullOrWhiteSpace(preview))
                message += " Response preview: " + preview;

            return message;
        }

        private static string TruncateForDiagnostics(string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;

            if (maxLength <= 0 || value.Length <= maxLength)
                return value;

            return value.Substring(0, maxLength) + "...";
        }


        //Комманды для парсинга ответа============================================================================================================

        /// <summary>
        /// Парсит Markdown-разметку и добавляет форматированный текст в RichTextBox
        /// </summary>
        private void ParseMarkdownToRichTextBox(RichTextBox rtb, string markdown)
        {
            if (string.IsNullOrEmpty(markdown))
                return;

            _diagnosticLogger.LogEvent(_currentDiagnosticRequestId, "ParseMarkdownToRichTextBox.START", new Dictionary<string, object>
            {
                { "length", markdown.Length },
                { "lines", markdown.Split('\n').Length }
            });
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
            _diagnosticLogger.LogEvent(_currentDiagnosticRequestId, "ParseMarkdownToRichTextBox.END", new Dictionary<string, object>
            {
                { "blocks", rtb.Document.Blocks.Count }
            });
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

            _diagnosticLogger.LogEvent(_currentDiagnosticRequestId, "ParseInlineMarkdown.START", new Dictionary<string, object>
            {
                { "length", text == null ? 0 : text.Length },
                { "preview", TrimForDiagnostics(text, 160) }
            });
            int pos = 0;
            int length = text.Length;
            int previousPos = 0;

            while (pos < length)
            {
                previousPos = pos;
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
                        else
                        {
                            paragraph.Inlines.Add(new Run(text.Substring(pos, 1)));
                            pos++;
                        }
                        break;
                }

                if (pos == previousPos)
                {
                    _diagnosticLogger.LogEvent(_currentDiagnosticRequestId, "ParseInlineMarkdown.NO_PROGRESS", new Dictionary<string, object>
                    {
                        { "position", pos },
                        { "markerType", markerType },
                        { "remainingPreview", TrimForDiagnostics(text.Substring(pos), 160) }
                    });
                    paragraph.Inlines.Add(new Run(text.Substring(pos, 1)));
                    pos++;
                }
            }
            _diagnosticLogger.LogEvent(_currentDiagnosticRequestId, "ParseInlineMarkdown.END", new Dictionary<string, object>
            {
                { "length", length }
            });
        }


    }
}


