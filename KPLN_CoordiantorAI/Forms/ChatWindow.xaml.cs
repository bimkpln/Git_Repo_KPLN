using KPLN_CoordiantorAI.Common;
using Autodesk.Revit.UI;
using KPLN_CoordiantorAI.ExternalModel;
using System;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Navigation;

namespace KPLN_CoordiantorAI.Forms
{
    public partial class ChatWindow : Window
    {
        private static readonly Regex SupportedHtmlTagRegex = new Regex(
            "<\\s*(?<closing>/)?\\s*(?<tag>a)\\b(?<attrs>[^>]*)>",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);
        private static readonly Regex HtmlHrefRegex = new Regex(
            "\\bhref\\s*=\\s*([\"'])(?<href>.*?)\\1",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);
        private static readonly Regex SourceLinkAttributeRegex = new Regex(
            "\\bdata-source\\s*=\\s*([\"'])(true|1)\\1",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);
        private static readonly Regex SourceLinkLabelRegex = new Regex(
            "\\bdata-source-label\\s*=\\s*([\"'])(?<label>.*?)\\1",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);
        private static readonly Regex SourceAnchorRegex = new Regex(
            "<\\s*a\\b(?<attrs>[^>]*)>.*?<\\s*/\\s*a\\s*>",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);
        private static readonly Regex MarkdownBoldRegex = new Regex(
            "\\*\\*(?<text>.+?)\\*\\*",
            RegexOptions.Singleline);
        private static readonly Regex HtmlLineBreakRegex = new Regex(
            "<\\s*br\\s*/?\\s*>|<\\s*/\\s*(p|div|li|ul|ol)\\s*>|<\\s*li\\b[^>]*>",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);
        private static readonly Regex AnyHtmlTagRegex = new Regex("<[^>]+>", RegexOptions.Singleline);
        private static readonly Regex BrokenHtmlTailRegex = new Regex(
            "<\\s*/?\\s*(a|b|strong|i|em|p|div|br|ul|ol|li)?\\s*$",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);
        private static readonly Regex MissingSentenceSpaceRegex = new Regex(
            "(?<=[.!?:;])(?=[A-ZА-ЯЁ])",
            RegexOptions.CultureInvariant);

        private readonly CoordinatorAiRepository _repository;
        private readonly IAiChatClient _aiClient;
        private readonly Bitrix24Client _bitrix24Client;
        private readonly GigaChatSettings _gigaChatSettings;
        private readonly Bitrix24Settings _bitrix24Settings;
        private readonly CurrentUserContext _userContext;
        private readonly ObservableCollection<ChatSession> _sessions;
        private readonly Autodesk.Revit.DB.Document _revitDocument;
        private readonly UIDocument _revitUiDocument;
        private ChatSession _session;
        private bool _isWaitingForAnswer;
        private bool _ignoreSessionSelectionChanged;
        private int _userMessageHistoryIndex = -1;
        private string _messageDraftBeforeHistory = string.Empty;

        public ChatWindow(
            CoordinatorAiRepository repository,
            IAiChatClient aiClient,
            CurrentUserContext userContext,
            Autodesk.Revit.DB.Document revitDocument,
            UIDocument revitUiDocument)
        {
            if (repository == null)
                throw new ArgumentNullException(nameof(repository));
            if (aiClient == null)
                throw new ArgumentNullException(nameof(aiClient));

            _repository = repository;
            _aiClient = aiClient;
            _bitrix24Client = new Bitrix24Client();
            _gigaChatSettings = _repository.LoadGigaChatSettings();
            _bitrix24Settings = _repository.LoadBitrix24Settings();
            _revitDocument = revitDocument;
            _revitUiDocument = revitUiDocument;
            _userContext = new CurrentUserContext
            {
                UserName = userContext == null ? Environment.UserName : userContext.UserName,
                SubDepartmentId = userContext == null ? -1 : userContext.SubDepartmentId
            };
            if (string.IsNullOrWhiteSpace(_userContext.UserName))
                _userContext.UserName = Environment.UserName;

            _sessions = new ObservableCollection<ChatSession>(_repository.LoadActiveChatSessionsForUser(_userContext.UserName));
            if (_sessions.Count == 0)
                _sessions.Add(CreateNewSession());

            foreach (ChatSession session in _sessions)
                PrepareSessionSourceLinks(session);

            _session = _sessions[0];

            InitializeComponent();
            InitializeModelTab();
            SessionsListBox.ItemsSource = _sessions;
            SessionsListBox.SelectedItem = _session;
            BindSession(_session);
            SetWaitingState(false);

            if (_session.Messages.Count == 0)
                SaveSessionSafely();
        }

        private ChatSession CreateNewSession()
        {
            return new ChatSession
            {
                UserName = _userContext.UserName,
                SubDepartmentId = _userContext.SubDepartmentId
            };
        }

        private void BindSession(ChatSession session)
        {
            if (session == null)
                return;

            _session = session;
            _ignoreSessionSelectionChanged = true;
            SessionsListBox.SelectedItem = _session;
            _ignoreSessionSelectionChanged = false;

            PrepareSessionSourceLinks(_session);
            DataContext = _session;
            ResetUserMessageHistoryNavigation();
            SessionInfoTextBlock.Text = string.Format(
                "{0}. Отдел {1}",
                string.IsNullOrWhiteSpace(_session.UserName) ? Environment.UserName : _session.UserName,
                SubDepartmentNameResolver.GetName(_session.SubDepartmentId));

            UpdateReactionButtons();
            ScrollMessagesToEnd();
            RefreshSessionsList();
        }

        private void RefreshSessionsList()
        {
            _ignoreSessionSelectionChanged = true;
            SessionsListBox.Items.Refresh();
            SessionsListBox.SelectedItem = _session;
            _ignoreSessionSelectionChanged = false;
        }

        private void SelectSession(ChatSession session)
        {
            if (session == null || ReferenceEquals(session, _session))
                return;

            SaveSessionSafely();
            BindSession(session);
        }

        private async void OnSendClick(object sender, RoutedEventArgs e)
        {
            await SendCurrentMessageAsync();
        }

        private async void OnMessageTextBoxPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Up && !HasTextInputModifier() && IsCaretOnFirstInputLine())
            {
                e.Handled = ShowPreviousUserMessage();
                return;
            }

            if (e.Key == Key.Down && !HasTextInputModifier() && IsCaretOnLastInputLine())
            {
                e.Handled = ShowNextUserMessage();
                return;
            }

            if (e.Key != Key.Enter || Keyboard.Modifiers.HasFlag(ModifierKeys.Shift))
                return;

            e.Handled = true;
            await SendCurrentMessageAsync();
        }

        private bool ShowPreviousUserMessage()
        {
            IList<string> history = GetCurrentUserMessageHistory();
            if (history.Count == 0)
                return false;

            if (_userMessageHistoryIndex < 0)
            {
                _messageDraftBeforeHistory = MessageTextBox.Text ?? string.Empty;
                _userMessageHistoryIndex = history.Count - 1;
            }
            else if (_userMessageHistoryIndex > 0)
            {
                _userMessageHistoryIndex--;
            }
            else
            {
                return true;
            }

            SetMessageTextFromHistory(history[_userMessageHistoryIndex]);
            return true;
        }

        private bool ShowNextUserMessage()
        {
            if (_userMessageHistoryIndex < 0)
                return false;

            IList<string> history = GetCurrentUserMessageHistory();
            if (history.Count == 0)
            {
                ResetUserMessageHistoryNavigation();
                return false;
            }

            if (_userMessageHistoryIndex < history.Count - 1)
            {
                _userMessageHistoryIndex++;
                SetMessageTextFromHistory(history[_userMessageHistoryIndex]);
            }
            else
            {
                string draft = _messageDraftBeforeHistory ?? string.Empty;
                ResetUserMessageHistoryNavigation();
                SetMessageTextFromHistory(draft);
            }

            return true;
        }

        private IList<string> GetCurrentUserMessageHistory()
        {
            List<string> history = new List<string>();
            if (_session == null || _session.Messages == null)
                return history;

            foreach (ChatMessage message in _session.Messages)
            {
                if (message == null || message.Role != ChatMessageRole.User || string.IsNullOrWhiteSpace(message.Text))
                    continue;

                history.Add(message.Text.Trim());
            }

            return history;
        }

        private void SetMessageTextFromHistory(string text)
        {
            MessageTextBox.Text = text ?? string.Empty;
            MessageTextBox.CaretIndex = MessageTextBox.Text.Length;
            MessageTextBox.ScrollToEnd();
        }

        private void ResetUserMessageHistoryNavigation()
        {
            _userMessageHistoryIndex = -1;
            _messageDraftBeforeHistory = string.Empty;
        }

        private static bool HasTextInputModifier()
        {
            return Keyboard.Modifiers != ModifierKeys.None;
        }

        private bool IsCaretOnFirstInputLine()
        {
            return MessageTextBox.GetLineIndexFromCharacterIndex(MessageTextBox.CaretIndex) <= 0;
        }

        private bool IsCaretOnLastInputLine()
        {
            int lineIndex = MessageTextBox.GetLineIndexFromCharacterIndex(MessageTextBox.CaretIndex);
            return lineIndex >= MessageTextBox.LineCount - 1;
        }

        private void OnNewChatClick(object sender, RoutedEventArgs e)
        {
            if (_isWaitingForAnswer)
                return;

            SaveSessionSafely();
            ChatSession newSession = CreateNewSession();
            _sessions.Insert(0, newSession);
            BindSession(newSession);
            SaveSessionSafely();
        }

        private void OnMainTabSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!ReferenceEquals(sender, MainTabControl))
                return;

            if (!_isWaitingForAnswer && RegulationsSearchTab != null && RegulationsSearchTab.IsSelected && MessageTextBox != null)
                MessageTextBox.Focus();
        }

        private void InitializeModelTab()
        {
            ModelTabContent.Content = CreateModelTabContent();
        }

        private FrameworkElement CreateModelTabContent()
        {
            if (_revitDocument == null || _revitUiDocument == null)
                return CreateModelUnavailableContent("Не найден активный документ Revit.");

            ExternalModelSettings settings = _repository.LoadExternalModelSettings();
            if (settings == null)
                settings = new ExternalModelSettings();

            ConnectionType connectionType;
            if (!Enum.TryParse(settings.ConnectionTypeName, true, out connectionType))
                return CreateModelUnavailableContent("В настройках не указан способ подключения внешней модели (conType).");

            if (connectionType == ConnectionType.OnlineAPI &&
                (string.IsNullOrWhiteSpace(settings.OnlineServerUrl) || string.IsNullOrWhiteSpace(settings.ApiKey)))
                return CreateModelUnavailableContent("Для OnlineAPI укажите ONLINE_SERVER_URL и API_KEY во вкладке \"Внешняя модель\".");

            if (connectionType == ConnectionType.LocalServer && string.IsNullOrWhiteSpace(settings.LocalServerUrl))
                return CreateModelUnavailableContent("Для LocalServer укажите LOCAL_SERVER_URL во вкладке \"Внешняя модель\".");

            return new ExternalModelControl(_revitDocument, _revitUiDocument, connectionType, settings);
        }

        private FrameworkElement CreateModelUnavailableContent(string message)
        {
            StackPanel panel = new StackPanel
            {
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                MaxWidth = 640
            };

            panel.Children.Add(new TextBlock
            {
                Text = "Работа с моделью недоступна",
                Foreground = Brushes.White,
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            panel.Children.Add(new TextBlock
            {
                Text = message,
                Foreground = CreateBrush(184, 192, 204),
                FontSize = 13,
                TextWrapping = TextWrapping.Wrap
            });

            return new Border
            {
                Background = CreateBrush(32, 36, 45),
                Padding = new Thickness(24),
                Child = panel
            };
        }

        private void OnDeleteChatClick(object sender, RoutedEventArgs e)
        {
            e.Handled = true;

            ChatSession sessionToDelete = (sender as FrameworkElement)?.Tag as ChatSession ?? _session;
            if (_isWaitingForAnswer || sessionToDelete == null)
                return;

            MessageBoxResult result = MessageBox.Show(
                this,
                "Вы хотите удалить чат?",
                "Координатор ИИ",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result != MessageBoxResult.Yes)
                return;

            try
            {
                bool wasCurrentSession = ReferenceEquals(sessionToDelete, _session);
                sessionToDelete.IsDeleted = true;
                _repository.MarkChatSessionDeleted(sessionToDelete.Id);
                _sessions.Remove(sessionToDelete);

                if (_sessions.Count == 0)
                {
                    ChatSession newSession = CreateNewSession();
                    _sessions.Add(newSession);
                }

                if (wasCurrentSession)
                {
                    BindSession(_sessions[0]);
                    SaveSessionSafely();
                }
                else
                {
                    RefreshSessionsList();
                }
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = "Ошибка удаления чата: " + ex.Message;
            }
        }

        private void OnChatCardMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            ChatSession selectedSession = (sender as FrameworkElement)?.Tag as ChatSession;
            if (_isWaitingForAnswer || selectedSession == null)
                return;

            SelectSession(selectedSession);
        }

        private void OnSessionsSelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (_ignoreSessionSelectionChanged)
                return;

            if (_isWaitingForAnswer)
            {
                SessionsListBox.SelectedItem = _session;
                return;
            }

            ChatSession selectedSession = SessionsListBox.SelectedItem as ChatSession;
            if (selectedSession == null || ReferenceEquals(selectedSession, _session))
                return;

            SelectSession(selectedSession);
        }

        private async Task SendCurrentMessageAsync()
        {
            if (_isWaitingForAnswer)
                return;

            string messageText = (MessageTextBox.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(messageText))
                return;

            ChatMessage userMessage = new ChatMessage
            {
                Role = ChatMessageRole.User,
                Text = messageText
            };

            _session.Messages.Add(userMessage);
            ResetUserMessageHistoryNavigation();
            MessageTextBox.Clear();
            ScrollMessagesToEnd();
            SaveSessionSafely();

            if (CoordinatorEscalationIntent.IsCoordinatorOfferRequest(messageText))
            {
                OfferCoordinatorHelp();
                return;
            }

            SetWaitingState(true);
            try
            {
                List<ChatMessage> requestMessages = _session.Messages.ToList();
                string answer = await _aiClient.SendAsync(requestMessages, _gigaChatSettings, CancellationToken.None);
                ChatMessage assistantMessage = new ChatMessage
                {
                    Role = ChatMessageRole.Assistant,
                    Text = answer,
                    CanRequestCoordinatorHelp = true
                };
                PrepareMessageSourceLinks(assistantMessage);
                _session.Messages.Add(assistantMessage);
            }
            catch (Exception ex)
            {
                AddCoordinatorOfferMessage("Ошибка получения ответа: " + ex.Message + "\r\nМожно подключить координатора по вашему отделу.");
            }
            finally
            {
                SetWaitingState(false);
                ScrollMessagesToEnd();
                SaveSessionSafely();
            }
        }

        private void OnLikeClick(object sender, RoutedEventArgs e)
        {
            SetReaction(_session.Reaction == ChatReaction.Like ? ChatReaction.None : ChatReaction.Like);
        }

        private void OnDislikeClick(object sender, RoutedEventArgs e)
        {
            ChatReaction newReaction = _session.Reaction == ChatReaction.Dislike ? ChatReaction.None : ChatReaction.Dislike;
            SetReaction(newReaction);

            if (newReaction != ChatReaction.Dislike)
                return;

            OfferCoordinatorHelp();
        }

        private void OnCoordinatorHelpClick(object sender, RoutedEventArgs e)
        {
            if (_isWaitingForAnswer)
                return;

            OfferCoordinatorHelp();
        }

        private async void OnCoordinatorCallClick(object sender, RoutedEventArgs e)
        {
            Bitrix24CoordinatorContact coordinator = (sender as FrameworkElement)?.Tag as Bitrix24CoordinatorContact;
            if (coordinator == null || _isWaitingForAnswer)
                return;

            MessageTextBox.IsEnabled = false;
            SendButton.IsEnabled = false;
            StatusTextBlock.Text = "Отправка сообщения координатору...";

            try
            {
                IList<Bitrix24CoordinatorContact> targets = GetCoordinatorCallTargets(coordinator);
                foreach (Bitrix24CoordinatorContact target in targets)
                    await _bitrix24Client.SendCoordinatorCallAsync(_bitrix24Settings, target, _session, CancellationToken.None);

                _session.Messages.Add(new ChatMessage
                {
                    Role = ChatMessageRole.Assistant,
                    Text = targets.Count > 1
                        ? "Сообщение отправлено всем ответственным координаторам отдела в Bitrix24."
                        : string.Format("Сообщение координатору {0} отправлено в Bitrix24.", coordinator.UserName)
                });
            }
            catch (Exception ex)
            {
                _session.Messages.Add(new ChatMessage
                {
                    Role = ChatMessageRole.Assistant,
                    Text = "Не удалось вызвать координатора через Bitrix24: " + ex.Message
                });
            }
            finally
            {
                MessageTextBox.IsEnabled = true;
                SendButton.IsEnabled = true;
                StatusTextBlock.Text = "Готово";
                ScrollMessagesToEnd();
                SaveSessionSafely();
                MessageTextBox.Focus();
            }
        }

        private IList<Bitrix24CoordinatorContact> GetCoordinatorCallTargets(Bitrix24CoordinatorContact selectedCoordinator)
        {
            List<Bitrix24CoordinatorContact> targets = new List<Bitrix24CoordinatorContact>();
            if (selectedCoordinator == null)
                return targets;

            foreach (Bitrix24DepartmentCoordinator department in _bitrix24Settings.DepartmentCoordinators)
            {
                if (department == null || department.DepartmentId != selectedCoordinator.DepartmentId)
                    continue;

                if (department.NotifyAllCoordinators)
                    targets.AddRange(department.GetConfiguredContacts());
                else
                    targets.Add(selectedCoordinator);

                break;
            }

            if (targets.Count == 0)
                targets.Add(selectedCoordinator);

            return targets;
        }

        private void AddCoordinatorOfferMessage(string text)
        {
            ChatMessage offerMessage = new ChatMessage
            {
                Role = ChatMessageRole.Assistant,
                Text = text
            };

            IList<Bitrix24CoordinatorContact> contacts = _bitrix24Settings.GetCoordinatorContacts(_session.SubDepartmentId);
            foreach (Bitrix24CoordinatorContact contact in contacts)
                offerMessage.CoordinatorOffers.Add(contact);

            if (offerMessage.CoordinatorOffers.Count == 0)
            {
                offerMessage.Text += "\r\nОтветственный координатор для отдела " +
                    SubDepartmentNameResolver.GetName(_session.SubDepartmentId) +
                    " пока не настроен.";
            }

            _session.Messages.Add(offerMessage);
            RefreshMessagesList();
        }

        private void RefreshMessagesList()
        {
            PrepareSessionSourceLinks(_session);
            if (MessagesListBox != null)
                MessagesListBox.Items.Refresh();
        }

        private void OnMessageRichTextBoxLoaded(object sender, RoutedEventArgs e)
        {
            RenderMessageHtml(sender as RichTextBox, (sender as RichTextBox)?.DataContext as ChatMessage);
        }

        private void OnMessageRichTextBoxDataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            RenderMessageHtml(sender as RichTextBox, e.NewValue as ChatMessage);
        }

        private void RenderMessageHtml(RichTextBox richTextBox, ChatMessage message)
        {
            if (richTextBox == null)
                return;

            Paragraph paragraph;
            richTextBox.Document = CreateMessageDocument(richTextBox, out paragraph);
            if (message == null)
                return;

            PrepareMessageSourceLinks(message);
            string html = RemoveSourceLinksFromHtml(message.Text ?? string.Empty);

            Stack<MessageInlineFrame> frames = new Stack<MessageInlineFrame>();
            Stack<MessageListFrame> lists = new Stack<MessageListFrame>();
            int textIndex = 0;
            foreach (Match match in SupportedHtmlTagRegex.Matches(html))
            {
                AddFormattedMessageText(GetMessageInlines(paragraph.Inlines, frames), html.Substring(textIndex, match.Index - textIndex));
                RenderMessageTag(
                    paragraph.Inlines,
                    frames,
                    lists,
                    match.Groups["tag"].Value.ToLowerInvariant(),
                    match.Groups["attrs"].Value,
                    match.Groups["closing"].Success);
                textIndex = match.Index + match.Length;
            }

            AddFormattedMessageText(GetMessageInlines(paragraph.Inlines, frames), html.Substring(textIndex));
        }

        private static FlowDocument CreateMessageDocument(RichTextBox richTextBox, out Paragraph paragraph)
        {
            FlowDocument document = new FlowDocument
            {
                PagePadding = new Thickness(0),
                Background = Brushes.Transparent,
                FontFamily = richTextBox.FontFamily,
                FontSize = richTextBox.FontSize,
                Foreground = richTextBox.Foreground
            };

            paragraph = new Paragraph
            {
                Margin = new Thickness(0)
            };
            document.Blocks.Add(paragraph);
            return document;
        }

        private void RenderMessageTag(
            InlineCollection rootInlines,
            Stack<MessageInlineFrame> frames,
            Stack<MessageListFrame> lists,
            string tag,
            string attributes,
            bool isClosing)
        {
            if (isClosing)
            {
                CloseMessageTag(rootInlines, frames, lists, tag);
                return;
            }

            InlineCollection inlines = GetMessageInlines(rootInlines, frames);
            if (tag == "br")
            {
                AddMessageBreak(inlines);
                return;
            }

            if (tag == "p" || tag == "div")
            {
                AddMessageBlockBreak(inlines);
                return;
            }

            if (tag == "ul" || tag == "ol")
            {
                AddMessageBlockBreak(inlines);
                lists.Push(new MessageListFrame(tag == "ol"));
                return;
            }

            if (tag == "li")
            {
                AddMessageBlockBreak(inlines);
                AddPlainMessageText(inlines, lists.Count > 0 ? lists.Peek().GetNextPrefix() : "- ");
                return;
            }

            if (tag == "b" || tag == "strong")
            {
                return;
            }

            if (tag == "i" || tag == "em")
            {
                return;
            }

            if (tag == "a")
                AddMessageLink(rootInlines, frames, attributes);
        }

        private void CloseMessageTag(
            InlineCollection rootInlines,
            Stack<MessageInlineFrame> frames,
            Stack<MessageListFrame> lists,
            string tag)
        {
            InlineCollection inlines = GetMessageInlines(rootInlines, frames);
            if (tag == "p" || tag == "div" || tag == "li")
            {
                AddMessageBlockBreak(inlines);
                return;
            }

            if (tag == "ul" || tag == "ol")
            {
                if (lists.Count > 0)
                    lists.Pop();

                AddMessageBlockBreak(inlines);
                return;
            }

            CloseMessageInlineFrame(frames, tag);
        }

        private void AddMessageLink(InlineCollection rootInlines, Stack<MessageInlineFrame> frames, string attributes)
        {
            Uri uri;
            if (!TryGetMessageLinkUri(attributes, out uri))
            {
                return;
            }

            if (SourceLinkAttributeRegex.IsMatch(attributes ?? string.Empty))
            {
                return;
            }

            Hyperlink hyperlink = new Hyperlink
            {
                NavigateUri = uri,
                Foreground = Brushes.LightSkyBlue
            };
            hyperlink.RequestNavigate += OnMessageHyperlinkRequestNavigate;
            GetMessageInlines(rootInlines, frames).Add(hyperlink);
            frames.Push(new MessageInlineFrame("a", hyperlink.Inlines));
        }

        private void PrepareSessionSourceLinks(ChatSession session)
        {
            if (session == null || session.Messages == null)
                return;

            foreach (ChatMessage message in session.Messages)
                PrepareMessageSourceLinks(message);
        }

        private void PrepareMessageSourceLinks(ChatMessage message)
        {
            if (message == null || message.SourceLinks == null)
                return;

            message.SourceLinks.Clear();
            string html = message.Text ?? string.Empty;
            if (string.IsNullOrWhiteSpace(html))
                return;

            HashSet<string> knownUrls = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Match match in SupportedHtmlTagRegex.Matches(html))
            {
                if (match.Groups["closing"].Success)
                    continue;

                string tag = match.Groups["tag"].Value;
                string attributes = match.Groups["attrs"].Value;
                if (!string.Equals(tag, "a", StringComparison.OrdinalIgnoreCase) ||
                    !SourceLinkAttributeRegex.IsMatch(attributes ?? string.Empty))
                {
                    continue;
                }

                Uri uri;
                if (!TryGetMessageLinkUri(attributes, out uri) || !knownUrls.Add(uri.AbsoluteUri))
                    continue;

                message.SourceLinks.Add(new MessageSourceLink
                {
                    Caption = GetSourceLinkLabel(attributes),
                    Url = uri.AbsoluteUri
                });
            }
        }

        private static string RemoveSourceLinksFromHtml(string html)
        {
            if (string.IsNullOrEmpty(html))
                return string.Empty;

            return SourceAnchorRegex.Replace(html, match =>
                SourceLinkAttributeRegex.IsMatch(match.Groups["attrs"].Value ?? string.Empty)
                    ? string.Empty
                    : match.Value);
        }

        private static bool TryGetMessageLinkUri(string attributes, out Uri uri)
        {
            uri = null;
            Match hrefMatch = HtmlHrefRegex.Match(attributes ?? string.Empty);
            string decodedHref = hrefMatch.Success
                ? WebUtility.HtmlDecode(hrefMatch.Groups["href"].Value).Trim()
                : string.Empty;

            return Uri.TryCreate(decodedHref, UriKind.Absolute, out uri) &&
                (string.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase));
        }

        private static string GetSourceLinkLabel(string attributes)
        {
            Match labelMatch = SourceLinkLabelRegex.Match(attributes ?? string.Empty);
            string label = labelMatch.Success
                ? WebUtility.HtmlDecode(labelMatch.Groups["label"].Value).Trim()
                : string.Empty;

            return string.IsNullOrWhiteSpace(label) ? "Источник" : label;
        }

        private static InlineCollection GetMessageInlines(InlineCollection rootInlines, Stack<MessageInlineFrame> frames)
        {
            return frames.Count == 0 ? rootInlines : frames.Peek().Inlines;
        }

        private static void CloseMessageInlineFrame(Stack<MessageInlineFrame> frames, string tag)
        {
            if (!frames.Any(frame => MessageTagsMatch(frame.Tag, tag)))
                return;

            while (frames.Count > 0)
            {
                MessageInlineFrame frame = frames.Pop();
                if (MessageTagsMatch(frame.Tag, tag))
                    return;
            }
        }

        private static bool MessageTagsMatch(string openTag, string closeTag)
        {
            if (string.Equals(openTag, closeTag, StringComparison.OrdinalIgnoreCase))
                return true;

            return (IsBoldTag(openTag) && IsBoldTag(closeTag)) ||
                (IsItalicTag(openTag) && IsItalicTag(closeTag));
        }

        private static bool IsBoldTag(string tag)
        {
            return tag == "b" || tag == "strong";
        }

        private static bool IsItalicTag(string tag)
        {
            return tag == "i" || tag == "em";
        }

        private static void AddFormattedMessageText(InlineCollection inlines, string html)
        {
            string text = GetMessageDisplayText(html);
            if (string.IsNullOrWhiteSpace(text))
                return;

            AddPlainMessageText(inlines, MarkdownBoldRegex.Replace(text, "${text}"));
        }

        private static void AddPlainMessageText(InlineCollection inlines, string text)
        {
            string normalized = (text ?? string.Empty).Replace("\r\n", "\n").Replace("\r", "\n");
            string[] lines = normalized.Split('\n');
            bool hasRenderedLine = false;
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                if (hasRenderedLine)
                {
                    AddMessageBreak(inlines);
                }
                else if (ShouldInsertSpaceBeforeText(inlines, line))
                {
                    inlines.Add(new Run(" "));
                }

                inlines.Add(new Run(line));
                hasRenderedLine = true;
            }
        }

        private static bool ShouldInsertSpaceBeforeText(InlineCollection inlines, string text)
        {
            char firstChar;
            char lastChar;
            return TryGetFirstNonWhiteSpaceChar(text, out firstChar) &&
                char.IsLetterOrDigit(firstChar) &&
                TryGetLastTextChar(inlines, out lastChar) &&
                IsSentenceJoinPunctuation(lastChar);
        }

        private static bool TryGetFirstNonWhiteSpaceChar(string text, out char value)
        {
            value = '\0';
            if (string.IsNullOrEmpty(text))
                return false;

            foreach (char currentChar in text)
            {
                if (char.IsWhiteSpace(currentChar))
                    continue;

                value = currentChar;
                return true;
            }

            return false;
        }

        private static bool TryGetLastTextChar(InlineCollection inlines, out char value)
        {
            value = '\0';
            if (inlines == null || inlines.Count == 0)
                return false;

            foreach (Inline inline in inlines.Cast<Inline>().Reverse())
            {
                if (TryGetLastTextChar(inline, out value))
                    return true;
            }

            return false;
        }

        private static bool TryGetLastTextChar(Inline inline, out char value)
        {
            value = '\0';
            Run run = inline as Run;
            if (run != null && !string.IsNullOrEmpty(run.Text))
            {
                for (int i = run.Text.Length - 1; i >= 0; i--)
                {
                    if (char.IsWhiteSpace(run.Text[i]))
                        continue;

                    value = run.Text[i];
                    return true;
                }
            }

            Span span = inline as Span;
            return span != null && TryGetLastTextChar(span.Inlines, out value);
        }

        private static bool IsSentenceJoinPunctuation(char value)
        {
            return value == '.' ||
                value == '!' ||
                value == '?' ||
                value == ':' ||
                value == ';';
        }

        private static void AddMessageBlockBreak(InlineCollection inlines)
        {
            if (inlines.Count > 0)
                AddMessageBreak(inlines);
        }

        private static void AddMessageBreak(InlineCollection inlines)
        {
            if (inlines.LastInline is LineBreak)
                return;

            inlines.Add(new LineBreak());
        }

        private static string GetMessageDisplayText(string html)
        {
            string text = HtmlLineBreakRegex.Replace(html ?? string.Empty, "\n");
            text = AnyHtmlTagRegex.Replace(text, string.Empty);
            text = WebUtility.HtmlDecode(text);
            text = AnyHtmlTagRegex.Replace(text, string.Empty);
            text = BrokenHtmlTailRegex.Replace(text, string.Empty);
            text = text.Replace('\u2013', '-').Replace('\u2014', '-');
            return MissingSentenceSpaceRegex.Replace(text, " ");
        }

        private void OnMessageHyperlinkRequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            if (e == null || e.Uri == null)
                return;

            OpenMessageUri(e.Uri);
            e.Handled = true;
        }

        private void OnSourceLinkButtonClick(object sender, RoutedEventArgs e)
        {
            Uri uri = (sender as FrameworkElement)?.Tag as Uri;
            if (uri == null)
                return;

            OpenMessageUri(uri);
            e.Handled = true;
        }

        private void OpenMessageUri(Uri uri)
        {
            try
            {
                Process.Start(new ProcessStartInfo(uri.AbsoluteUri)
                {
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = "Не удалось открыть ссылку: " + ex.Message;
            }
        }

        private sealed class MessageInlineFrame
        {
            public MessageInlineFrame(string tag, InlineCollection inlines)
            {
                Tag = tag;
                Inlines = inlines;
            }

            public string Tag { get; private set; }

            public InlineCollection Inlines { get; private set; }
        }

        private sealed class MessageListFrame
        {
            private int _itemNumber;

            public MessageListFrame(bool isOrdered)
            {
                IsOrdered = isOrdered;
            }

            public bool IsOrdered { get; private set; }

            public string GetNextPrefix()
            {
                if (!IsOrdered)
                    return "- ";

                _itemNumber++;
                return _itemNumber + ". ";
            }
        }

        private void OfferCoordinatorHelp()
        {
            AddCoordinatorOfferMessage("Можно подключить координатора по вашему отделу.");
            ScrollMessagesToEnd();
            SaveSessionSafely();
        }

        private void SetReaction(ChatReaction reaction)
        {
            _session.Reaction = reaction;
            UpdateReactionButtons();
            SaveSessionSafely();
        }

        private void UpdateReactionButtons()
        {
            Brush likeDefaultBackground = CreateBrush(65, 108, 88);
            Brush dislikeDefaultBackground = CreateBrush(89, 34, 42);
            Brush likeBackground = CreateBrush(37, 150, 92);
            Brush dislikeBackground = CreateBrush(176, 70, 70);
            Brush likeSelectedBorder = CreateBrush(128, 226, 170);
            Brush dislikeSelectedBorder = CreateBrush(230, 150, 150);

            LikeButton.Background = _session.Reaction == ChatReaction.Like ? likeBackground : likeDefaultBackground;
            DislikeButton.Background = _session.Reaction == ChatReaction.Dislike ? dislikeBackground : dislikeDefaultBackground;
            LikeButton.BorderBrush = _session.Reaction == ChatReaction.Like ? likeSelectedBorder : likeDefaultBackground;
            DislikeButton.BorderBrush = _session.Reaction == ChatReaction.Dislike ? dislikeSelectedBorder : dislikeDefaultBackground;
            LikeButton.BorderThickness = _session.Reaction == ChatReaction.Like ? new Thickness(2) : new Thickness(1);
            DislikeButton.BorderThickness = _session.Reaction == ChatReaction.Dislike ? new Thickness(2) : new Thickness(1);
        }

        private void OnLikeButtonMouseEnter(object sender, MouseEventArgs e)
        {
            Brush hoverBackground = _session.Reaction == ChatReaction.Like
                ? CreateBrush(37, 150, 92)
                : CreateBrush(67, 142, 107);
            Brush hoverBorder = _session.Reaction == ChatReaction.Like
                ? CreateBrush(128, 226, 170)
                : CreateBrush(0, 117, 63);

            LikeButton.Background = hoverBackground;
            LikeButton.BorderBrush = hoverBorder;
        }

        private void OnDislikeButtonMouseEnter(object sender, MouseEventArgs e)
        {
            Brush hoverBackground = _session.Reaction == ChatReaction.Dislike
                ? CreateBrush(176, 70, 70)
                : CreateBrush(172, 53, 70);
            Brush hoverBorder = _session.Reaction == ChatReaction.Dislike
                ? CreateBrush(230, 150, 150)
                : CreateBrush(160, 13, 34);

            DislikeButton.Background = hoverBackground;
            DislikeButton.BorderBrush = hoverBorder;
        }

        private void OnReactionButtonMouseLeave(object sender, MouseEventArgs e)
        {
            UpdateReactionButtons();
        }

        private void OnNewChatButtonMouseEnter(object sender, MouseEventArgs e)
        {
            Brush hoverBackground = CreateBrush(74, 131, 244);
            NewChatButton.Background = hoverBackground;
            NewChatButton.BorderBrush = hoverBackground;
        }

        private void OnNewChatButtonMouseLeave(object sender, MouseEventArgs e)
        {
            Brush background = CreateBrush(44, 107, 237);
            NewChatButton.Background = background;
            NewChatButton.BorderBrush = background;
        }

        private static Brush CreateBrush(byte red, byte green, byte blue)
        {
            return new SolidColorBrush(Color.FromRgb(red, green, blue));
        }

        private void SetWaitingState(bool isWaiting)
        {
            _isWaitingForAnswer = isWaiting;
            MessageTextBox.IsEnabled = !isWaiting;
            SendButton.IsEnabled = !isWaiting;
            SessionsListBox.IsEnabled = !isWaiting;
            NewChatButton.IsEnabled = !isWaiting;
            StatusTextBlock.Text = isWaiting ? "Ожидание ответа ИИ..." : "Готово";

            if (!isWaiting && RegulationsSearchTab != null && RegulationsSearchTab.IsSelected)
                MessageTextBox.Focus();
        }

        private void ScrollMessagesToEnd()
        {
            if (_session.Messages.Count == 0)
                return;

            MessagesListBox.ScrollIntoView(_session.Messages[_session.Messages.Count - 1]);
        }

        private void SaveSessionSafely()
        {
            if (_session == null || _session.IsDeleted)
                return;

            try
            {
                _repository.SaveChatSession(_session);
                SessionsListBox.Items.Refresh();
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = "Ошибка сохранения БД: " + ex.Message;
            }
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            SaveSessionSafely();
            base.OnClosing(e);
        }
    }
}
