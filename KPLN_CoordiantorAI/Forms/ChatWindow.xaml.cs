using KPLN_CoordiantorAI.Common;
using System;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace KPLN_CoordiantorAI.Forms
{
    public partial class ChatWindow : Window
    {
        private readonly CoordinatorAiRepository _repository;
        private readonly IAiChatClient _aiClient;
        private readonly Bitrix24Client _bitrix24Client;
        private readonly GigaChatSettings _gigaChatSettings;
        private readonly Bitrix24Settings _bitrix24Settings;
        private readonly CurrentUserContext _userContext;
        private readonly ObservableCollection<ChatSession> _sessions;
        private ChatSession _session;
        private bool _isWaitingForAnswer;
        private bool _ignoreSessionSelectionChanged;

        public ChatWindow(CoordinatorAiRepository repository, IAiChatClient aiClient, CurrentUserContext userContext)
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

            _session = _sessions[0];

            InitializeComponent();
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

            DataContext = _session;
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
            if (e.Key != Key.Enter || Keyboard.Modifiers.HasFlag(ModifierKeys.Shift))
                return;

            e.Handled = true;
            await SendCurrentMessageAsync();
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
            MessageTextBox.Clear();
            ScrollMessagesToEnd();
            SaveSessionSafely();

            if (IsCoordinatorCallRequest(messageText))
            {
                AddCoordinatorOfferMessage("Можно подключить координатора по вашему отделу.");
                ScrollMessagesToEnd();
                SaveSessionSafely();
                return;
            }

            SetWaitingState(true);
            try
            {
                List<ChatMessage> requestMessages = _session.Messages.ToList();
                string answer = await _aiClient.SendAsync(requestMessages, _gigaChatSettings, CancellationToken.None);
                _session.Messages.Add(new ChatMessage
                {
                    Role = ChatMessageRole.Assistant,
                    Text = answer
                });
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

            AddCoordinatorOfferMessage("Можно подключить координатора по вашему отделу.");
            RefreshMessagesList();
            ScrollMessagesToEnd();
            SaveSessionSafely();
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
            if (MessagesListBox != null)
                MessagesListBox.Items.Refresh();
        }

        private static bool IsCoordinatorCallRequest(string messageText)
        {
            string text = (messageText ?? string.Empty).ToLowerInvariant();
            if (text.IndexOf("координатор", StringComparison.Ordinal) < 0)
                return false;

            return text.IndexOf("выз", StringComparison.Ordinal) >= 0 ||
                text.IndexOf("поз", StringComparison.Ordinal) >= 0 ||
                text.IndexOf("зов", StringComparison.Ordinal) >= 0 ||
                text.IndexOf("подключ", StringComparison.Ordinal) >= 0 ||
                text.IndexOf("нужен", StringComparison.Ordinal) >= 0 ||
                text.IndexOf("нужна", StringComparison.Ordinal) >= 0;
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

            if (!isWaiting)
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