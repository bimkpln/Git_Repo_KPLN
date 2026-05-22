using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace KPLN_CoordiantorAI.Common
{
    public sealed class Bitrix24Client
    {
        public async Task SendCoordinatorCallAsync(
            Bitrix24Settings settings,
            Bitrix24CoordinatorContact coordinator,
            ChatSession session,
            CancellationToken cancellationToken)
        {
            if (settings == null || string.IsNullOrWhiteSpace(settings.WebhookUrl))
                throw new InvalidOperationException("В настройках Bitrix24 не заполнен Webhook URL.");

            if (coordinator == null || string.IsNullOrWhiteSpace(coordinator.UserId))
                throw new InvalidOperationException("Для отдела не указан ответственный Bitrix24.");

            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;

            using (HttpClient httpClient = new HttpClient())
            {
                httpClient.Timeout = TimeSpan.FromSeconds(30);
                string messageText = await BuildCoordinatorCallMessageAsync(settings, session, coordinator, httpClient, cancellationToken).ConfigureAwait(false);

                using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, BuildMethodUrl(settings.WebhookUrl, "im.message.add")))
                {
                    request.Content = new FormUrlEncodedContent(new Dictionary<string, string>
                    {
                        { "DIALOG_ID", coordinator.UserId.Trim() },
                        { "MESSAGE", messageText },
                        { "SYSTEM", "N" },
                        { "URL_PREVIEW", "N" }
                    });

                    HttpResponseMessage response;
                    try
                    {
                        response = await httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false);
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException("Не удалось отправить сообщение в Bitrix24. " + ex.Message, ex);
                    }

                    string responseBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new InvalidOperationException(
                            string.Format("Bitrix24 вернул HTTP {0} {1}. {2}", (int)response.StatusCode, response.ReasonPhrase, responseBody));
                    }

                    if (!string.IsNullOrWhiteSpace(responseBody) &&
                        responseBody.IndexOf("\"error\"", StringComparison.OrdinalIgnoreCase) >= 0)
                        throw new InvalidOperationException("Bitrix24 вернул ошибку. " + responseBody.Trim());
                }
            }
        }

        private static string BuildMethodUrl(string webhookUrl, string methodName)
        {
            string normalizedUrl = webhookUrl.Trim().TrimEnd('/');
            Uri uri;
            if (Uri.TryCreate(normalizedUrl, UriKind.Absolute, out uri))
            {
                string[] segments = uri.AbsolutePath.Trim('/').Split('/');
                if (segments.Length >= 3 && string.Equals(segments[0], "rest", StringComparison.OrdinalIgnoreCase))
                {
                    string basePath = string.Format("/rest/{0}/{1}", segments[1], segments[2]);
                    UriBuilder builder = new UriBuilder(uri)
                    {
                        Path = basePath + "/" + methodName,
                        Query = string.Empty
                    };

                    return builder.Uri.ToString().TrimEnd('/');
                }
            }

            if (normalizedUrl.EndsWith("/" + methodName, StringComparison.OrdinalIgnoreCase) ||
                normalizedUrl.EndsWith("/" + methodName + ".json", StringComparison.OrdinalIgnoreCase))
                return normalizedUrl;

            return normalizedUrl + "/" + methodName;
        }

        private static async Task<string> BuildCoordinatorCallMessageAsync(
            Bitrix24Settings settings,
            ChatSession session,
            Bitrix24CoordinatorContact coordinator,
            HttpClient httpClient,
            CancellationToken cancellationToken)
        {
            string userName = session == null || string.IsNullOrWhiteSpace(session.UserName)
                ? Environment.UserName
                : session.UserName;
            string departmentName = coordinator == null
                ? string.Empty
                : coordinator.DepartmentName;
            string userCaption = await BuildUserCaptionAsync(settings, userName, httpClient, cancellationToken).ConfigureAwait(false);
            string payload = GetCallPayload(settings, session);
            bool hasQuestion = !IsNoQuestionPayload(payload);

            return string.Format(
                hasQuestion
                    ? "Пользователь {0} просит подключиться координатора по отделу [b]{1}[/b] в приложении [b]\"Координатор ИИ\"[/b] для помощи в решении вопроса. {2}"
                    : "Пользователь {0} просит подключиться координатора по отделу [b]{1}[/b] в приложении [b]\"Координатор ИИ\"[/b].  {2}",
                userCaption,
                EscapeBitrixText(departmentName),
                payload);
        }

        private static string GetCallPayload(Bitrix24Settings settings, ChatSession session)
        {
            if (session == null || session.Messages == null)
                return "Координатор был призван без вопросов :)";

            string firstQuestion = GetFirstQuestionText(session);
            if (string.IsNullOrWhiteSpace(firstQuestion))
                return "Координатор был призван без вопросов :)";

            if (settings != null && settings.CoordinatorMessageMode == Bitrix24CoordinatorMessageMode.FirstQuestion)
                return "[u]Вопрос следующий:[/u] \n" + firstQuestion;

            string transcript = ChatTranscriptFormatter.BuildTranscript(session.Messages);
            return string.IsNullOrWhiteSpace(transcript) ? "[u]Вопрос следующий:[/u] \n" + firstQuestion : "[u]Вопрос следующий:[/u] \n" + transcript;
        }

        private static bool IsNoQuestionPayload(string payload)
        {
            return string.Equals(
                payload,
                "Координатор был призван без вопросов :)",
                StringComparison.Ordinal);
        }

        private static string GetFirstQuestionText(ChatSession session)
        {
            foreach (ChatMessage message in session.Messages)
            {
                if (message == null ||
                    message.Role != ChatMessageRole.User ||
                    string.IsNullOrWhiteSpace(message.Text) ||
                    CoordinatorEscalationIntent.IsCoordinatorOfferRequest(message.Text))
                    continue;

                return message.Text.Trim();
            }

            return string.Empty;
        }

        private static async Task<string> BuildUserCaptionAsync(
            Bitrix24Settings settings,
            string userName,
            HttpClient httpClient,
            CancellationToken cancellationToken)
        {
            string displayName = string.IsNullOrWhiteSpace(userName) ? Environment.UserName : userName.Trim();
            Bitrix24User user = await FindBitrixUserAsync(settings, displayName, httpClient, cancellationToken).ConfigureAwait(false);
            if (user != null && !string.IsNullOrWhiteSpace(user.Id))
                return string.Format("[b][user={0}]{1}[/user][/b]", user.Id.Trim(), EscapeBitrixText(displayName));

            return "[b]" + EscapeBitrixText(displayName) + "[/b]";
        }

        private static async Task<Bitrix24User> FindBitrixUserAsync(
            Bitrix24Settings settings,
            string displayName,
            HttpClient httpClient,
            CancellationToken cancellationToken)
        {
            if (settings == null || string.IsNullOrWhiteSpace(settings.WebhookUrl) || string.IsNullOrWhiteSpace(displayName))
                return null;

            foreach (string query in BuildUserSearchQueries(displayName))
            {
                Bitrix24User user = await SearchBitrixUserAsync(settings, query, displayName, httpClient, cancellationToken).ConfigureAwait(false);
                if (user != null)
                    return user;
            }

            return null;
        }

        private static IEnumerable<string> BuildUserSearchQueries(string displayName)
        {
            yield return displayName;

            string[] parts = displayName.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length >= 2)
                yield return parts[1] + " " + parts[0];
        }

        private static async Task<Bitrix24User> SearchBitrixUserAsync(
            Bitrix24Settings settings,
            string query,
            string displayName,
            HttpClient httpClient,
            CancellationToken cancellationToken)
        {
            try
            {
                using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, BuildMethodUrl(settings.WebhookUrl, "user.search")))
                {
                    request.Content = new FormUrlEncodedContent(new Dictionary<string, string>
                    {
                        { "FILTER[FIND]", query }
                    });

                    HttpResponseMessage response = await httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false);
                    string responseBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode ||
                        string.IsNullOrWhiteSpace(responseBody) ||
                        responseBody.IndexOf("\"error\"", StringComparison.OrdinalIgnoreCase) >= 0)
                        return null;

                    Bitrix24UserSearchResponse searchResponse = Deserialize<Bitrix24UserSearchResponse>(responseBody);
                    if (searchResponse == null || searchResponse.Result == null || searchResponse.Result.Count == 0)
                        return null;

                    Bitrix24User exactUser = FindExactUser(searchResponse.Result, displayName);
                    if (exactUser != null)
                        return exactUser;

                    return searchResponse.Result.Count == 1 ? searchResponse.Result[0] : null;
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static Bitrix24User FindExactUser(IList<Bitrix24User> users, string displayName)
        {
            string normalizedDisplayName = NormalizeName(displayName);
            foreach (Bitrix24User user in users)
            {
                string nameLastName = NormalizeName((user.Name ?? string.Empty) + " " + (user.LastName ?? string.Empty));
                string lastNameName = NormalizeName((user.LastName ?? string.Empty) + " " + (user.Name ?? string.Empty));
                if (string.Equals(normalizedDisplayName, nameLastName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(normalizedDisplayName, lastNameName, StringComparison.OrdinalIgnoreCase))
                    return user;
            }

            return null;
        }

        private static string NormalizeName(string value)
        {
            return string.Join(" ", (value ?? string.Empty).Trim().ToLowerInvariant().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));
        }

        private static string EscapeBitrixText(string value)
        {
            return (value ?? string.Empty).Replace("[", "(").Replace("]", ")");
        }

        private static T Deserialize<T>(string json)
        {
            DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(T));
            using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(json ?? string.Empty)))
            {
                return (T)serializer.ReadObject(stream);
            }
        }

        [DataContract]
        private sealed class Bitrix24UserSearchResponse
        {
            [DataMember(Name = "result")]
            public List<Bitrix24User> Result { get; set; }
        }

        [DataContract]
        private sealed class Bitrix24User
        {
            [DataMember(Name = "ID")]
            public string Id { get; set; }

            [DataMember(Name = "NAME")]
            public string Name { get; set; }

            [DataMember(Name = "LAST_NAME")]
            public string LastName { get; set; }
        }
    }
}