using System;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace KPLN_CoordiantorAI.Common
{
    public enum ChatMessageRole
    {
        User = 0,
        Assistant = 1
    }

    public enum ChatReaction
    {
        Dislike = -1,
        None = 0,
        Like = 1
    }

    public enum Bitrix24CoordinatorMessageMode
    {
        FirstQuestion = 0,
        FullChat = 1
    }

    public sealed class ChatMessage
    {
        public ChatMessage()
        {
            Id = Guid.NewGuid().ToString("N");
            CreatedAt = DateTime.Now;
            CoordinatorOffers = new ObservableCollection<Bitrix24CoordinatorContact>();
        }

        public string Id { get; set; }

        public ChatMessageRole Role { get; set; }

        public string Text { get; set; }

        public DateTime CreatedAt { get; set; }

        public ObservableCollection<Bitrix24CoordinatorContact> CoordinatorOffers { get; private set; }

        public bool IsUserMessage
        {
            get { return Role == ChatMessageRole.User; }
        }

        public bool HasCoordinatorOffers
        {
            get { return CoordinatorOffers != null && CoordinatorOffers.Count > 0; }
        }

        public bool CanRequestCoordinatorHelp { get; set; }

        public string RoleCaption
        {
            get { return IsUserMessage ? "Пользователь" : "ИИ"; }
        }

        public string CreatedAtCaption
        {
            get { return CreatedAt.ToString("HH:mm"); }
        }
    }

    public sealed class ChatSession
    {
        public ChatSession()
        {
            Id = Guid.NewGuid().ToString("N");
            CreatedAt = DateTime.Now;
            UpdatedAt = CreatedAt;
            Reaction = ChatReaction.None;
            Messages = new ObservableCollection<ChatMessage>();
        }

        public string Id { get; set; }

        public DateTime CreatedAt { get; set; }

        public DateTime UpdatedAt { get; set; }

        public string UserName { get; set; }

        public int SubDepartmentId { get; set; }

        public ChatReaction Reaction { get; set; }

        public bool IsDeleted { get; set; }

        public ObservableCollection<ChatMessage> Messages { get; private set; }

        public string Title
        {
            get
            {
                foreach (ChatMessage message in Messages)
                {
                    if (message != null && message.Role == ChatMessageRole.User && !string.IsNullOrWhiteSpace(message.Text))
                        return TrimForCaption(message.Text, 34);
                }

                return "Новый чат";
            }
        }

        public string FullQuestion
        {
            get
            {
                foreach (ChatMessage message in Messages)
                {
                    if (message != null && message.Role == ChatMessageRole.User && !string.IsNullOrWhiteSpace(message.Text))
                        return message.Text.Trim();
                }

                return "Новый чат";
            }
        }

        public string UpdatedAtCaption
        {
            get { return UpdatedAt.ToString("dd.MM HH:mm"); }
        }

        private static string TrimForCaption(string value, int maxLength)
        {
            string normalized = (value ?? string.Empty).Replace("\r", " ").Replace("\n", " ").Trim();
            if (normalized.Length <= maxLength)
                return normalized;

            return normalized.Substring(0, maxLength).TrimEnd() + "...";
        }
    }

    public sealed class CurrentUserContext
    {
        public string UserName { get; set; }

        public int SubDepartmentId { get; set; }
    }

    public sealed class GigaChatSettings
    {
        public GigaChatSettings()
        {
            AuthUrl = "https://ngw.devices.sberbank.ru:9443/api/v2/oauth";
            ApiUrl = "https://gigachat.devices.sberbank.ru/api/v1";
            Scope = "GIGACHAT_API_PERS";
            SystemPrompt =
                "Ты ИИ-помощник KPLN CoordinatorAI. Отвечай на русском языке.\r\n" +
                "Стиль общения: сдержанный, спокойный и профессиональный. Пиши по делу, без лишней воды.\r\n" +
                "Иногда можно использовать легкий уместный юмор, но без фамильярности и без шуток в серьезных ситуациях.\r\n" +
                "Если точного ответа нет или данных недостаточно, прямо скажи об этом. Не выдумывай факты, ссылки, нормы, версии программ и чужие решения.\r\n" +
                "Если вопрос неоднозначный, задай короткий уточняющий вопрос или явно перечисли допущения.\r\n" +
                "Когда даешь инструкцию, делай ее пошаговой и проверяемой.";
        }

        public string AuthUrl { get; set; }

        public string ApiUrl { get; set; }

        public string ClientId { get; set; }

        public string ClientSecret { get; set; }

        public string Scope { get; set; }

        public string CertificatePath { get; set; }

        public string EmbeddingFolderPath { get; set; }

        public string SystemPrompt { get; set; }
    }

    public sealed class SubDepartmentInfo
    {
        public int Id { get; set; }

        public string Name { get; set; }
    }

    public sealed class Bitrix24Settings
    {
        public Bitrix24Settings()
        {
            CoordinatorMessageMode = Bitrix24CoordinatorMessageMode.FullChat;
            DepartmentCoordinators = new ObservableCollection<Bitrix24DepartmentCoordinator>();
        }

        public string WebhookUrl { get; set; }

        public Bitrix24CoordinatorMessageMode CoordinatorMessageMode { get; set; }

        public ObservableCollection<Bitrix24DepartmentCoordinator> DepartmentCoordinators { get; private set; }

        public IList<Bitrix24CoordinatorContact> GetCoordinatorContacts(int subDepartmentId)
        {
            List<Bitrix24CoordinatorContact> contacts = new List<Bitrix24CoordinatorContact>();
            foreach (Bitrix24DepartmentCoordinator coordinator in DepartmentCoordinators)
            {
                if (coordinator == null || coordinator.DepartmentId != subDepartmentId)
                    continue;

                contacts.AddRange(coordinator.GetConfiguredContacts());
                break;
            }

            return contacts;
        }
    }

    public sealed class Bitrix24DepartmentCoordinator
    {
        public Bitrix24DepartmentCoordinator()
        {
            Coordinators = new ObservableCollection<Bitrix24CoordinatorContact>();
        }

        public int DepartmentId { get; set; }

        public string DepartmentName { get; set; }

        public bool NotifyAllCoordinators { get; set; }

        public ObservableCollection<Bitrix24CoordinatorContact> Coordinators { get; private set; }

        public IList<Bitrix24CoordinatorContact> GetConfiguredContacts()
        {
            List<Bitrix24CoordinatorContact> contacts = new List<Bitrix24CoordinatorContact>();
            foreach (Bitrix24CoordinatorContact contact in Coordinators)
            {
                if (contact != null && !string.IsNullOrWhiteSpace(contact.UserId))
                    contacts.Add(contact);
            }

            return contacts;
        }
    }

    public sealed class Bitrix24CoordinatorContact
    {
        public int DepartmentId { get; set; }

        public string DepartmentName { get; set; }

        public string UserId { get; set; }

        public string UserName { get; set; }

        public string ButtonCaption
        {
            get { return string.Format("Написать в Битрикс24 координатору {0} ({1})", DepartmentName, UserName); }
        }

        public string ChipCaption
        {
            get { return string.Format("{0} ({1})", UserName, UserId); }
        }
    }

    public static class CoordinatorEscalationIntent
    {
        private static readonly string[] CoordinatorCallMarkers =
        {
            "выз",
            "поз",
            "зов",
            "подключ",
            "нужен",
            "нужна"
        };

        private static readonly string[] AnswerFeedbackMarkers =
        {
            "ответ невер",
            "ответ не вер",
            "неверный ответ",
            "не верный ответ",
            "ответ неправиль",
            "ответ не правиль",
            "неправильный ответ",
            "не правильный ответ",
            "ответ не помог",
            "ответ не тот",
            "не тот ответ",
            "ошибка в ответе",
            "ошибочный ответ",
            "это неверно",
            "это не верно",
            "это неправильно",
            "это не правильно",
            "это не так"
        };

        private static readonly string[] ShortAnswerFeedbackMarkers =
        {
            "неверно",
            "не верно",
            "неправильно",
            "не правильно",
            "не помог",
            "не помогло"
        };

        public static bool IsCoordinatorOfferRequest(string messageText)
        {
            return IsCoordinatorCallRequest(messageText) || IsUnhelpfulAnswerFeedback(messageText);
        }

        public static bool IsCoordinatorCallRequest(string messageText)
        {
            string text = Normalize(messageText);
            if (text.IndexOf("координатор", StringComparison.Ordinal) < 0)
                return false;

            return ContainsAny(text, CoordinatorCallMarkers);
        }

        private static bool IsUnhelpfulAnswerFeedback(string messageText)
        {
            string text = Normalize(messageText);
            if (string.IsNullOrWhiteSpace(text))
                return false;

            if (ContainsAny(text, AnswerFeedbackMarkers))
                return true;

            return text.Length <= 80 && ContainsAny(text, ShortAnswerFeedbackMarkers);
        }

        private static string Normalize(string value)
        {
            return (value ?? string.Empty).Trim().ToLowerInvariant().Replace('ё', 'е');
        }

        private static bool ContainsAny(string text, IEnumerable<string> markers)
        {
            foreach (string marker in markers)
            {
                if (text.IndexOf(marker, StringComparison.Ordinal) >= 0)
                    return true;
            }

            return false;
        }
    }

    public static class ChatTranscriptFormatter
    {
        public static string BuildTranscript(IEnumerable<ChatMessage> messages)
        {
            if (messages == null)
                return string.Empty;

            StringBuilder builder = new StringBuilder();
            foreach (ChatMessage message in messages)
            {
                if (message == null)
                    continue;

                builder
                    .Append('[')
                    .Append(message.CreatedAt.ToString("yyyy-MM-dd HH:mm:ss"))
                    .Append("] ")
                    .Append(message.RoleCaption)
                    .Append(": ")
                    .AppendLine(message.Text ?? string.Empty);
            }

            return builder.ToString().TrimEnd();
        }

        public static IList<ChatMessage> ParseTranscript(string transcript)
        {
            List<ChatMessage> messages = new List<ChatMessage>();
            if (string.IsNullOrWhiteSpace(transcript))
                return messages;

            ChatMessage currentMessage = null;
            StringBuilder currentText = null;
            string[] lines = transcript.Replace("\r\n", "\n").Split('\n');
            foreach (string line in lines)
            {
                DateTime createdAt;
                ChatMessageRole role;
                string text;
                if (TryParseTranscriptHeader(line, out createdAt, out role, out text))
                {
                    AddParsedMessage(messages, currentMessage, currentText);
                    currentMessage = new ChatMessage
                    {
                        CreatedAt = createdAt,
                        Role = role
                    };
                    currentText = new StringBuilder(text ?? string.Empty);
                    continue;
                }

                if (currentMessage == null)
                    continue;

                currentText.AppendLine();
                currentText.Append(line);
            }

            AddParsedMessage(messages, currentMessage, currentText);
            return messages;
        }

        private static void AddParsedMessage(ICollection<ChatMessage> messages, ChatMessage message, StringBuilder text)
        {
            if (message == null)
                return;

            message.Text = text == null ? string.Empty : text.ToString();
            messages.Add(message);
        }

        private static bool TryParseTranscriptHeader(
            string line,
            out DateTime createdAt,
            out ChatMessageRole role,
            out string text)
        {
            createdAt = DateTime.Now;
            role = ChatMessageRole.Assistant;
            text = string.Empty;

            if (string.IsNullOrEmpty(line) || line[0] != '[')
                return false;

            int endDateIndex = line.IndexOf("] ", StringComparison.Ordinal);
            if (endDateIndex <= 1)
                return false;

            string dateText = line.Substring(1, endDateIndex - 1);
            if (!DateTime.TryParseExact(
                dateText,
                "yyyy-MM-dd HH:mm:ss",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out createdAt))
                return false;

            string messagePart = line.Substring(endDateIndex + 2);
            int roleSeparatorIndex = messagePart.IndexOf(": ", StringComparison.Ordinal);
            if (roleSeparatorIndex <= 0)
                return false;

            string roleText = messagePart.Substring(0, roleSeparatorIndex);
            if (string.Equals(roleText, "Пользователь", StringComparison.OrdinalIgnoreCase))
                role = ChatMessageRole.User;
            else if (string.Equals(roleText, "ИИ", StringComparison.OrdinalIgnoreCase))
                role = ChatMessageRole.Assistant;
            else
                return false;

            text = messagePart.Substring(roleSeparatorIndex + 2);
            return true;
        }
    }
}