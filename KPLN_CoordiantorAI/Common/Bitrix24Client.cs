using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
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
            using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, BuildMethodUrl(settings.WebhookUrl, "im.message.add")))
            {
                httpClient.Timeout = TimeSpan.FromSeconds(30);
                request.Content = new FormUrlEncodedContent(new Dictionary<string, string>
                {
                    { "DIALOG_ID", coordinator.UserId.Trim() },
                    { "MESSAGE", BuildCoordinatorCallMessage(settings, session, coordinator) },
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

        private static string BuildCoordinatorCallMessage(
            Bitrix24Settings settings,
            ChatSession session,
            Bitrix24CoordinatorContact coordinator)
        {
            string userName = session == null || string.IsNullOrWhiteSpace(session.UserName)
                ? Environment.UserName
                : session.UserName;
            string departmentName = coordinator == null
                ? string.Empty
                : coordinator.DepartmentName;
            string chatPayload = GetChatPayload(settings, session);

            return string.Format(
                "Пользователь {0} просит подключить координатора по отделу {1}.\n\n{2}",
                userName,
                departmentName,
                string.IsNullOrWhiteSpace(chatPayload) ? "Пока нет сообщений." : chatPayload);
        }

        private static string GetChatPayload(Bitrix24Settings settings, ChatSession session)
        {
            if (session == null || session.Messages == null)
                return string.Empty;

            if (settings != null && settings.CoordinatorMessageMode == Bitrix24CoordinatorMessageMode.FirstQuestion)
            {
                foreach (ChatMessage message in session.Messages)
                {
                    if (message != null && message.Role == ChatMessageRole.User && !string.IsNullOrWhiteSpace(message.Text))
                        return "Первое сообщение:\n" + message.Text.Trim();
                }

                return string.Empty;
            }

            string transcript = ChatTranscriptFormatter.BuildTranscript(session.Messages);
            return string.IsNullOrWhiteSpace(transcript) ? string.Empty : "Чат:\n" + transcript;
        }
    }
}