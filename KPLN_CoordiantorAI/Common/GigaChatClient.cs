using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Security;
using System.Security.Authentication;
using System.Security.Cryptography.X509Certificates;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace KPLN_CoordiantorAI.Common
{
    public sealed class GigaChatClient : IAiChatClient
    {
        private const string CertificateErrorMessage = "Сертификат GigaChat либо не найден, либо недействителен.";

        private string _accessToken;
        private DateTime _accessTokenExpiresAtUtc;

        public async Task<string> SendAsync(IEnumerable<ChatMessage> messages, GigaChatSettings settings, CancellationToken cancellationToken)
        {
            ValidateSettings(settings);

            string token = await GetAccessTokenAsync(settings, cancellationToken).ConfigureAwait(false);
            GigaChatCompletionResponse response = await GetCompletionAsync(messages, settings, token, cancellationToken).ConfigureAwait(false);

            if (response == null || response.Choices == null || response.Choices.Count == 0 || response.Choices[0].Message == null)
                throw new InvalidOperationException("GigaChat вернул пустой ответ без choices[0].message.");

            string content = response.Choices[0].Message.Content;
            if (string.IsNullOrWhiteSpace(content))
                throw new InvalidOperationException("GigaChat вернул пустой текст ответа.");

            return content.Trim();
        }

        private async Task<string> GetAccessTokenAsync(GigaChatSettings settings, CancellationToken cancellationToken)
        {
            if (!string.IsNullOrWhiteSpace(_accessToken) && _accessTokenExpiresAtUtc > DateTime.UtcNow.AddMinutes(1))
                return _accessToken;

            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;

            using (HttpClient httpClient = CreateHttpClient(settings, TimeSpan.FromSeconds(60)))
            using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, settings.AuthUrl.Trim()))
            {
                request.Headers.TryAddWithoutValidation("RqUID", Guid.NewGuid().ToString());
                request.Headers.TryAddWithoutValidation("Authorization", "Basic " + BuildAuthorizationKey(settings));
                request.Headers.TryAddWithoutValidation("Accept", "application/json");
                request.Content = new StringContent(
                    "scope=" + Uri.EscapeDataString(settings.Scope.Trim()),
                    Encoding.UTF8,
                    "application/x-www-form-urlencoded");

                HttpResponseMessage response;
                try
                {
                    response = await httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("Не удалось выполнить запрос токена GigaChat. " + BuildExceptionDetails(ex), ex);
                }

                string responseBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (!response.IsSuccessStatusCode)
                    throw new InvalidOperationException(BuildHttpErrorMessage("получения токена", response, responseBody));

                GigaChatTokenResponse tokenResponse = Deserialize<GigaChatTokenResponse>(responseBody);
                if (tokenResponse == null || string.IsNullOrWhiteSpace(tokenResponse.AccessToken))
                    throw new InvalidOperationException("GigaChat не вернул access_token.");

                _accessToken = tokenResponse.AccessToken;
                _accessTokenExpiresAtUtc = GetTokenExpiresAtUtc(tokenResponse);
                return _accessToken;
            }
        }

        private async Task<GigaChatCompletionResponse> GetCompletionAsync(
            IEnumerable<ChatMessage> messages,
            GigaChatSettings settings,
            string token,
            CancellationToken cancellationToken)
        {
            using (HttpClient httpClient = CreateHttpClient(settings, TimeSpan.FromSeconds(120)))
            using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, GetChatCompletionsUrl(settings)))
            {
                request.Headers.TryAddWithoutValidation("Authorization", "Bearer " + token);
                request.Headers.TryAddWithoutValidation("Accept", "application/json");
                request.Content = new StringContent(
                    Serialize(BuildCompletionRequest(messages, settings)),
                    Encoding.UTF8,
                    "application/json");

                HttpResponseMessage response;
                try
                {
                    response = await httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("Не удалось выполнить запрос ответа GigaChat. " + BuildExceptionDetails(ex), ex);
                }

                string responseBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (!response.IsSuccessStatusCode)
                    throw new InvalidOperationException(BuildHttpErrorMessage("получения ответа", response, responseBody));

                return Deserialize<GigaChatCompletionResponse>(responseBody);
            }
        }

        private static void ValidateSettings(GigaChatSettings settings)
        {
            if (settings == null)
                throw new InvalidOperationException("Настройки GigaChat не загрузились из БД.");

            List<string> missingFields = new List<string>();
            if (string.IsNullOrWhiteSpace(settings.AuthUrl))
                missingFields.Add("Auth URL");
            if (string.IsNullOrWhiteSpace(settings.ApiUrl))
                missingFields.Add("API URL");
            if (string.IsNullOrWhiteSpace(settings.ClientId))
                missingFields.Add("Client ID");
            if (string.IsNullOrWhiteSpace(settings.ClientSecret))
                missingFields.Add("Client Secret");
            if (string.IsNullOrWhiteSpace(settings.Scope))
                missingFields.Add("Scope");

            if (missingFields.Count > 0)
                throw new InvalidOperationException("В настройках GigaChat не заполнены поля: " + string.Join(", ", missingFields));
        }

        private static HttpClient CreateHttpClient(GigaChatSettings settings, TimeSpan timeout)
        {
            HttpClientHandler handler = new HttpClientHandler();
            X509Certificate2 trustedCertificate = LoadTrustedCertificate(settings.CertificatePath);
            if (trustedCertificate != null)
            {
                handler.ServerCertificateCustomValidationCallback = (request, certificate, chain, sslPolicyErrors) =>
                    ValidateServerCertificate(certificate, chain, sslPolicyErrors, trustedCertificate);
            }

            HttpClient httpClient = new HttpClient(handler);
            httpClient.Timeout = timeout;
            return httpClient;
        }

        private static X509Certificate2 LoadTrustedCertificate(string certificatePath)
        {
            if (string.IsNullOrWhiteSpace(certificatePath))
                return null;

            string normalizedPath = certificatePath.Trim().Trim('"');
            if (!File.Exists(normalizedPath))
                throw new InvalidOperationException(CertificateErrorMessage);

            try
            {
                if (string.Equals(Path.GetExtension(normalizedPath), ".pem", StringComparison.OrdinalIgnoreCase))
                    return new X509Certificate2(ExtractPemCertificateBytes(normalizedPath));

                return new X509Certificate2(normalizedPath);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(CertificateErrorMessage, ex);
            }
        }

        private static byte[] ExtractPemCertificateBytes(string certificatePath)
        {
            const string beginMarker = "-----BEGIN CERTIFICATE-----";
            const string endMarker = "-----END CERTIFICATE-----";

            string pem = File.ReadAllText(certificatePath);
            int beginIndex = pem.IndexOf(beginMarker, StringComparison.Ordinal);
            int endIndex = pem.IndexOf(endMarker, StringComparison.Ordinal);
            if (beginIndex < 0 || endIndex < 0 || endIndex <= beginIndex)
                throw new InvalidOperationException(CertificateErrorMessage);

            beginIndex += beginMarker.Length;
            string base64 = pem.Substring(beginIndex, endIndex - beginIndex)
                .Replace("\r", string.Empty)
                .Replace("\n", string.Empty)
                .Trim();

            return Convert.FromBase64String(base64);
        }

        private static bool ValidateServerCertificate(
            X509Certificate2 serverCertificate,
            X509Chain chain,
            SslPolicyErrors sslPolicyErrors,
            X509Certificate2 trustedCertificate)
        {
            if (sslPolicyErrors == SslPolicyErrors.None)
                return true;

            if (serverCertificate == null || trustedCertificate == null)
                return false;

            if ((sslPolicyErrors & SslPolicyErrors.RemoteCertificateNameMismatch) == SslPolicyErrors.RemoteCertificateNameMismatch)
                return false;

            if ((sslPolicyErrors & SslPolicyErrors.RemoteCertificateNotAvailable) == SslPolicyErrors.RemoteCertificateNotAvailable)
                return false;

            if (DateTime.Now < serverCertificate.NotBefore || DateTime.Now > serverCertificate.NotAfter)
                return false;

            if (CertificateThumbprintsEqual(serverCertificate, trustedCertificate))
                return true;

            using (X509Chain customChain = new X509Chain())
            {
                customChain.ChainPolicy.ExtraStore.Add(trustedCertificate);
                customChain.ChainPolicy.RevocationMode = X509RevocationMode.NoCheck;
                customChain.ChainPolicy.VerificationFlags = X509VerificationFlags.AllowUnknownCertificateAuthority;

                return customChain.Build(serverCertificate) && ChainContainsCertificate(customChain, trustedCertificate);
            }
        }

        private static bool ChainContainsCertificate(X509Chain chain, X509Certificate2 certificate)
        {
            if (chain == null || certificate == null)
                return false;

            foreach (X509ChainElement element in chain.ChainElements)
            {
                if (CertificateThumbprintsEqual(element.Certificate, certificate))
                    return true;
            }

            return false;
        }

        private static bool CertificateThumbprintsEqual(X509Certificate2 first, X509Certificate2 second)
        {
            return first != null &&
                second != null &&
                string.Equals(first.Thumbprint, second.Thumbprint, StringComparison.OrdinalIgnoreCase);
        }

        private static string BuildAuthorizationKey(GigaChatSettings settings)
        {
            string credentials = string.Format("{0}:{1}", settings.ClientId.Trim(), settings.ClientSecret.Trim());
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(credentials));
        }

        private static string GetChatCompletionsUrl(GigaChatSettings settings)
        {
            string apiUrl = settings.ApiUrl.Trim().TrimEnd('/');
            if (apiUrl.EndsWith("/chat/completions", StringComparison.OrdinalIgnoreCase))
                return apiUrl;

            return apiUrl + "/chat/completions";
        }

        private static GigaChatCompletionRequest BuildCompletionRequest(IEnumerable<ChatMessage> messages, GigaChatSettings settings)
        {
            GigaChatCompletionRequest request = new GigaChatCompletionRequest
            {
                Model = "GigaChat",
                Messages = new List<GigaChatMessage>()
            };

            if (!string.IsNullOrWhiteSpace(settings.SystemPrompt))
            {
                request.Messages.Add(new GigaChatMessage
                {
                    Role = "system",
                    Content = settings.SystemPrompt
                });
            }

            if (messages != null)
            {
                foreach (ChatMessage message in messages.Where(m => m != null && !string.IsNullOrWhiteSpace(m.Text)))
                {
                    request.Messages.Add(new GigaChatMessage
                    {
                        Role = message.Role == ChatMessageRole.User ? "user" : "assistant",
                        Content = message.Text
                    });
                }
            }

            return request;
        }

        private static DateTime GetTokenExpiresAtUtc(GigaChatTokenResponse tokenResponse)
        {
            if (tokenResponse.ExpiresAt > 0)
            {
                long expiresAt = tokenResponse.ExpiresAt;
                if (expiresAt > 9999999999)
                    return UnixEpochUtc.AddMilliseconds(expiresAt);

                return UnixEpochUtc.AddSeconds(expiresAt);
            }

            return DateTime.UtcNow.AddMinutes(25);
        }

        private static DateTime UnixEpochUtc
        {
            get { return new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc); }
        }

        private static string BuildHttpErrorMessage(string operationName, HttpResponseMessage response, string responseBody)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("Ошибка GigaChat при ")
                .Append(operationName)
                .Append(": HTTP ")
                .Append((int)response.StatusCode)
                .Append(" ")
                .Append(response.ReasonPhrase);

            if (!string.IsNullOrWhiteSpace(responseBody))
            {
                builder.AppendLine();
                builder.Append(responseBody.Trim());
            }

            return builder.ToString();
        }

        private static string BuildExceptionDetails(Exception exception)
        {
            if (ContainsAuthenticationException(exception))
            {
                return CertificateErrorMessage;
            }

            StringBuilder builder = new StringBuilder();
            Exception currentException = exception;
            while (currentException != null)
            {
                if (builder.Length > 0)
                    builder.Append(" ");

                builder.Append(currentException.GetType().Name)
                    .Append(": ")
                    .Append(currentException.Message);

                currentException = currentException.InnerException;
            }

            return builder.ToString();
        }

        private static bool ContainsAuthenticationException(Exception exception)
        {
            Exception currentException = exception;
            while (currentException != null)
            {
                if (currentException is AuthenticationException)
                    return true;

                currentException = currentException.InnerException;
            }

            return false;
        }

        private static string Serialize<T>(T value)
        {
            DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(T));
            using (MemoryStream stream = new MemoryStream())
            {
                serializer.WriteObject(stream, value);
                return Encoding.UTF8.GetString(stream.ToArray());
            }
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
        private sealed class GigaChatTokenResponse
        {
            [DataMember(Name = "access_token")]
            public string AccessToken { get; set; }

            [DataMember(Name = "expires_at")]
            public long ExpiresAt { get; set; }
        }

        [DataContract]
        private sealed class GigaChatCompletionRequest
        {
            [DataMember(Name = "model")]
            public string Model { get; set; }

            [DataMember(Name = "messages")]
            public List<GigaChatMessage> Messages { get; set; }
        }

        [DataContract]
        private sealed class GigaChatMessage
        {
            [DataMember(Name = "role")]
            public string Role { get; set; }

            [DataMember(Name = "content")]
            public string Content { get; set; }
        }

        [DataContract]
        private sealed class GigaChatCompletionResponse
        {
            [DataMember(Name = "choices")]
            public List<GigaChatChoice> Choices { get; set; }
        }

        [DataContract]
        private sealed class GigaChatChoice
        {
            [DataMember(Name = "message")]
            public GigaChatMessage Message { get; set; }
        }
    }
}