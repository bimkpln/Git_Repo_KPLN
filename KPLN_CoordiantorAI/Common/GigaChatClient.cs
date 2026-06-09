using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Security;
using System.Security.Authentication;
using System.Security.Cryptography.X509Certificates;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace KPLN_CoordiantorAI.Common
{
    public sealed class GigaChatClient : IAiChatClient
    {
        private const string CertificateErrorMessage = "Сертификат GigaChat либо не найден, либо недействителен.";
        private static readonly Regex UsedSourcesRegex = new Regex(
            "\\[{2,}\\s*USED_SOURCES\\s*:\\s*(?<ids>[^\\]]*)\\]{2,}",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        private static readonly Regex SourceParentheticalMentionRegex = new Regex(
            "\\s*[\\(\\[]\\s*(?:(?:источник(?:е|а|у|ом)?|source)\\s*)?S\\d+\\s*[\\)\\]]",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        private static readonly Regex SourceAttributionMentionRegex = new Regex(
            "\\s*,?\\s*(?:привед[её]нн?ым(?:и)?|указанн?ым(?:и)?|описанн?ым(?:и)?)\\s+в\\s+источник(?:е|ах)?\\s*S\\d+",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        private static readonly Regex SourceAccordingMentionRegex = new Regex(
            "\\b(?:согласно|по)\\s+(?:источник(?:у|е)?\\s*)?S\\d+\\s*,?\\s*",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        private static readonly Regex SourceStandaloneMentionRegex = new Regex(
            "\\s*\\b(?:источник(?:е|а|у|ом)?|source)\\s*S\\d+\\b\\s*",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        private static readonly Regex HtmlAnchorTagRegex = new Regex(
            "<\\s*a\\b(?<attrs>[^>]*)>",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);
        private static readonly Regex HtmlHrefAttributeRegex = new Regex(
            "\\bhref\\s*=\\s*(?:(['\"])(?<href>.*?)\\1|(?<href>[^\\s>]+))",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);
        private static readonly Regex PlainHttpUrlRegex = new Regex(
            "https?://[^\\s\"'<>]+",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);

        private string _accessToken;
        private DateTime _accessTokenExpiresAtUtc;

        public async Task<string> SendAsync(IEnumerable<ChatMessage> messages, GigaChatSettings settings, CancellationToken cancellationToken)
        {
            ValidateSettings(settings);
            AiSearchOptions aiSearchOptions = AiSearchOptions.FromJson(settings.AiSearchSettingsJson);

            string token = await GetAccessTokenAsync(settings, cancellationToken).ConfigureAwait(false);
            string query = GetLastUserMessageText(messages);
            IList<MoodleEmbeddingSearchResult> searchResults = await GetMoodleArticlesAsync(messages, settings, aiSearchOptions, token, cancellationToken).ConfigureAwait(false);
            bool shouldUseArticleHintPrompt =
                aiSearchOptions.EnableEmbeddingSearch &&
                searchResults.Count == 0 &&
                !string.IsNullOrWhiteSpace(query) &&
                !string.IsNullOrWhiteSpace(settings.ArticleHintPrompt);
            if (!shouldUseArticleHintPrompt &&
                aiSearchOptions.EnableEmbeddingSearch &&
                searchResults.Count == 0 &&
                !aiSearchOptions.AllowAnswerWithoutContext &&
                !string.IsNullOrWhiteSpace(query))
            {
                return "Я не нашел подходящий источник в базе знаний Moodle для этого вопроса.";
            }

            GigaChatCompletionResponse response = await GetCompletionAsync(messages, searchResults, settings, aiSearchOptions, shouldUseArticleHintPrompt, token, cancellationToken).ConfigureAwait(false);

            if (response == null || response.Choices == null || response.Choices.Count == 0 || response.Choices[0].Message == null)
                throw new InvalidOperationException("GigaChat вернул пустой ответ без choices[0].message.");

            string content = response.Choices[0].Message.Content;
            if (string.IsNullOrWhiteSpace(content))
                throw new InvalidOperationException("GigaChat вернул пустой текст ответа.");

            return BuildAnswerWithUsedSources(content.Trim(), searchResults, aiSearchOptions);
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
            IList<MoodleEmbeddingSearchResult> searchResults,
            GigaChatSettings settings,
            AiSearchOptions aiSearchOptions,
            bool includeArticleHintPrompt,
            string token,
            CancellationToken cancellationToken)
        {
            using (HttpClient httpClient = CreateHttpClient(settings, TimeSpan.FromSeconds(120)))
            using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, GetChatCompletionsUrl(settings)))
            {
                request.Headers.TryAddWithoutValidation("Authorization", "Bearer " + token);
                request.Headers.TryAddWithoutValidation("Accept", "application/json");
                request.Content = new StringContent(
                    Serialize(BuildCompletionRequest(messages, searchResults, settings, aiSearchOptions, includeArticleHintPrompt)),
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

        private async Task<IList<MoodleEmbeddingSearchResult>> GetMoodleArticlesAsync(
            IEnumerable<ChatMessage> messages,
            GigaChatSettings settings,
            AiSearchOptions aiSearchOptions,
            string token,
            CancellationToken cancellationToken)
        {
            List<MoodleEmbeddingSearchResult> emptyResults = new List<MoodleEmbeddingSearchResult>();
            if (aiSearchOptions == null || !aiSearchOptions.EnableEmbeddingSearch)
                return emptyResults;

            MoodleEmbeddingIndex index = MoodleEmbeddingIndex.Load(settings.EmbeddingFolderPath, aiSearchOptions.MoodleUrlTemplate);
            string query = GetLastUserMessageText(messages);
            if (index == null || string.IsNullOrWhiteSpace(query))
                return emptyResults;

            List<MoodleEmbeddingSearchResult> candidates = new List<MoodleEmbeddingSearchResult>();
            foreach (MoodleEmbeddingSearchResult aliasResult in GetAliasSearchResults(index, settings, query))
                candidates.Add(aliasResult);

            IList<double> queryVector = await GetEmbeddingAsync(query, settings, aiSearchOptions, token, cancellationToken).ConfigureAwait(false);
            int candidateCount = Math.Min(aiSearchOptions.TopK * 5, 50);
            foreach (MoodleEmbeddingSearchResult embeddingResult in index.FindTop(queryVector, candidateCount, aiSearchOptions.MinSimilarity, query))
                candidates.Add(embeddingResult);

            return GetUniqueSourceResults(candidates, aiSearchOptions.TopK);
        }

        private static IList<MoodleEmbeddingSearchResult> GetAliasSearchResults(
            MoodleEmbeddingIndex index,
            GigaChatSettings settings,
            string query)
        {
            List<MoodleEmbeddingSearchResult> emptyResults = new List<MoodleEmbeddingSearchResult>();
            if (index == null || settings == null || string.IsNullOrWhiteSpace(query))
                return emptyResults;

            IList<ArticleAliasEntry> aliases;
            try
            {
                aliases = ArticleAliasSettings.FromJson(settings.ArticleAliasesJson);
            }
            catch
            {
                return emptyResults;
            }

            IList<string> matchedArticleIds = ArticleAliasSettings.FindMatchingArticleIds(query, aliases);
            return matchedArticleIds.Count == 0
                ? emptyResults
                : index.FindByArticleIds(matchedArticleIds, query);
        }

        private async Task<IList<double>> GetEmbeddingAsync(
            string text,
            GigaChatSettings settings,
            AiSearchOptions aiSearchOptions,
            string token,
            CancellationToken cancellationToken)
        {
            using (HttpClient httpClient = CreateHttpClient(settings, TimeSpan.FromSeconds(120)))
            using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, GetEmbeddingsUrl(settings)))
            {
                request.Headers.TryAddWithoutValidation("Authorization", "Bearer " + token);
                request.Headers.TryAddWithoutValidation("Accept", "application/json");
                request.Headers.TryAddWithoutValidation("RqUID", Guid.NewGuid().ToString());
                request.Content = new StringContent(
                    Serialize(new GigaChatEmbeddingRequest
                    {
                        Model = aiSearchOptions.EmbeddingModel,
                        Input = text
                    }),
                    Encoding.UTF8,
                    "application/json");

                HttpResponseMessage response;
                try
                {
                    response = await httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("Не удалось выполнить запрос embedding GigaChat. " + BuildExceptionDetails(ex), ex);
                }

                string responseBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (!response.IsSuccessStatusCode)
                    throw new InvalidOperationException(BuildHttpErrorMessage("получения embedding", response, responseBody));

                GigaChatEmbeddingResponse embeddingResponse = Deserialize<GigaChatEmbeddingResponse>(responseBody);
                if (embeddingResponse == null ||
                    embeddingResponse.Data == null ||
                    embeddingResponse.Data.Count == 0 ||
                    embeddingResponse.Data[0].Embedding == null ||
                    embeddingResponse.Data[0].Embedding.Count == 0)
                {
                    throw new InvalidOperationException("GigaChat вернул пустой embedding для вопроса.");
                }

                return embeddingResponse.Data[0].Embedding;
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

        private static string GetEmbeddingsUrl(GigaChatSettings settings)
        {
            string apiUrl = settings.ApiUrl.Trim().TrimEnd('/');
            if (apiUrl.EndsWith("/embeddings", StringComparison.OrdinalIgnoreCase))
                return apiUrl;

            return apiUrl + "/embeddings";
        }

        private static GigaChatCompletionRequest BuildCompletionRequest(
            IEnumerable<ChatMessage> messages,
            IList<MoodleEmbeddingSearchResult> searchResults,
            GigaChatSettings settings,
            AiSearchOptions aiSearchOptions,
            bool includeArticleHintPrompt)
        {
            GigaChatCompletionRequest request = new GigaChatCompletionRequest
            {
                Model = aiSearchOptions.ChatModel,
                Messages = new List<GigaChatMessage>()
            };

            string systemMessage = BuildSystemMessage(settings, searchResults, aiSearchOptions, includeArticleHintPrompt);
            if (!string.IsNullOrWhiteSpace(systemMessage))
            {
                request.Messages.Add(new GigaChatMessage
                {
                    Role = "system",
                    Content = systemMessage
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

        private static string BuildSystemMessage(
            GigaChatSettings settings,
            IList<MoodleEmbeddingSearchResult> searchResults,
            AiSearchOptions aiSearchOptions,
            bool includeArticleHintPrompt)
        {
            StringBuilder builder = new StringBuilder();
            if (settings != null && !string.IsNullOrWhiteSpace(settings.SystemPrompt))
                builder.AppendLine(settings.SystemPrompt.Trim());

            if (searchResults != null && searchResults.Count > 0)
            {
                int remainingChars = aiSearchOptions == null ? 16000 : aiSearchOptions.MaxContextChars;
                int articleIndex = 0;
                foreach (MoodleEmbeddingSearchResult result in searchResults)
                {
                    if (result == null || result.Article == null || remainingChars <= 0)
                        continue;

                    articleIndex++;
                    string articleBody = GetArticleBodyContext(result.Article);
                    if (articleBody.Length > remainingChars)
                        articleBody = articleBody.Substring(0, remainingChars);

                    builder.AppendLine();
                    string sourceId = GetSourceId(articleIndex);
                    builder.AppendLine("moodle_context_block:");
                    AppendArticleContext(builder, result, sourceId, aiSearchOptions);
                    if (aiSearchOptions != null && aiSearchOptions.DebugRetrieval)
                    {
                        builder.Append("score: ")
                            .Append(result.Score.ToString("0.000", CultureInfo.InvariantCulture))
                            .AppendLine();
                    }

                    if (!string.IsNullOrWhiteSpace(articleBody))
                    {
                        builder.AppendLine("article_body:");
                        builder.Append(articleBody);
                        remainingChars -= articleBody.Length;
                    }
                }
            }

            if (settings != null && !string.IsNullOrWhiteSpace(settings.ResponseContextPrompt))
            {
                if (builder.Length > 0)
                    builder.AppendLine();

                builder.AppendLine(settings.ResponseContextPrompt.Trim());
            }

            if (includeArticleHintPrompt && settings != null && !string.IsNullOrWhiteSpace(settings.ArticleHintPrompt))
            {
                if (builder.Length > 0)
                    builder.AppendLine();

                builder.AppendLine(settings.ArticleHintPrompt.Trim());
            }

            return builder.ToString().Trim();
        }

        private static string GetArticleBodyContext(MoodleEmbeddingArticle article)
        {
            if (article == null)
                return string.Empty;

            if (!string.IsNullOrWhiteSpace(article.Html))
                return article.Html.Trim();

            return string.IsNullOrWhiteSpace(article.Text) ? string.Empty : article.Text.Trim();
        }

        private static void AppendArticleContext(
            StringBuilder builder,
            MoodleEmbeddingSearchResult result,
            string sourceId,
            AiSearchOptions aiSearchOptions)
        {
            MoodleEmbeddingArticle article = result.Article;
            builder.Append("marker_id: ").Append(sourceId).AppendLine();
            if (!string.IsNullOrWhiteSpace(article.ArticleId))
                builder.Append("article_id: ").Append(article.ArticleId.Trim()).AppendLine();
            if (!string.IsNullOrWhiteSpace(article.Title))
                builder.Append("title: ").Append(article.Title.Trim()).AppendLine();
            if (!string.IsNullOrWhiteSpace(article.SourceUrl))
                builder.Append("url: ").Append(article.SourceUrl.Trim()).AppendLine();
            if (!string.IsNullOrWhiteSpace(article.GetHeadingsText()))
                builder.Append("headings: ").Append(article.GetHeadingsText()).AppendLine();
            if (result.RelevantSections != null && result.RelevantSections.Count > 0)
            {
                builder.AppendLine("relevant_sections:");
                foreach (MoodleEmbeddingSection section in result.RelevantSections)
                {
                    if (section == null || string.IsNullOrWhiteSpace(section.Text))
                        continue;

                    builder.Append("- heading_path: ").Append(section.GetHeadingPathText()).AppendLine();
                    builder.Append("  text: ").Append(section.Text.Trim()).AppendLine();
                }
            }
        }

        private static string GetLastUserMessageText(IEnumerable<ChatMessage> messages)
        {
            ChatMessage message = messages == null
                ? null
                : messages.LastOrDefault(m => m != null && m.Role == ChatMessageRole.User && !string.IsNullOrWhiteSpace(m.Text));

            return message == null ? string.Empty : message.Text.Trim();
        }

        private static string BuildAnswerWithUsedSources(
            string answer,
            IList<MoodleEmbeddingSearchResult> searchResults,
            AiSearchOptions aiSearchOptions)
        {
            UsedSourceMarker marker = ExtractUsedSourceMarker(answer);
            IList<MoodleEmbeddingSearchResult> usedResults = marker.Found
                ? GetUsedSourceResults(searchResults, marker.SourceIds)
                : GetPrimarySourceResults(searchResults);
            if (usedResults.Count == 0)
                usedResults = GetPrimarySourceResults(searchResults);

            return AppendSourceLinks(CleanVisibleSourceIdMentions(marker.CleanAnswer), usedResults, aiSearchOptions);
        }

        private static string CleanVisibleSourceIdMentions(string answer)
        {
            string cleanAnswer = answer ?? string.Empty;
            cleanAnswer = SourceParentheticalMentionRegex.Replace(cleanAnswer, string.Empty);
            cleanAnswer = SourceAttributionMentionRegex.Replace(cleanAnswer, string.Empty);
            cleanAnswer = SourceAccordingMentionRegex.Replace(cleanAnswer, string.Empty);
            cleanAnswer = SourceStandaloneMentionRegex.Replace(cleanAnswer, string.Empty);
            cleanAnswer = Regex.Replace(cleanAnswer, "[ \\t]+:", ":", RegexOptions.CultureInvariant);
            cleanAnswer = Regex.Replace(cleanAnswer, "[ \\t]{2,}", " ", RegexOptions.CultureInvariant);
            return cleanAnswer.Trim();
        }

        private static UsedSourceMarker ExtractUsedSourceMarker(string answer)
        {
            string value = answer ?? string.Empty;
            Match match = UsedSourcesRegex.Match(value);
            if (!match.Success)
                return new UsedSourceMarker(value.Trim(), new List<string>(), false);

            List<string> sourceIds = new List<string>();
            string[] rawIds = match.Groups["ids"].Value.Split(',');
            foreach (string rawId in rawIds)
            {
                string id = (rawId ?? string.Empty).Trim().ToUpperInvariant();
                if (id.Length == 0 || sourceIds.Contains(id))
                    continue;

                sourceIds.Add(id);
            }

            string cleanAnswer = UsedSourcesRegex.Replace(value, string.Empty).Trim();
            return new UsedSourceMarker(cleanAnswer, sourceIds, true);
        }

        private static IList<MoodleEmbeddingSearchResult> GetUsedSourceResults(
            IList<MoodleEmbeddingSearchResult> searchResults,
            IList<string> sourceIds)
        {
            List<MoodleEmbeddingSearchResult> usedResults = new List<MoodleEmbeddingSearchResult>();
            if (searchResults == null || sourceIds == null || sourceIds.Count == 0)
                return usedResults;

            for (int i = 0; i < searchResults.Count; i++)
            {
                if (sourceIds.Contains(GetSourceId(i + 1)))
                    usedResults.Add(searchResults[i]);
            }

            return usedResults;
        }

        private static IList<MoodleEmbeddingSearchResult> GetPrimarySourceResults(IList<MoodleEmbeddingSearchResult> searchResults)
        {
            return GetUniqueSourceResults(searchResults, 1);
        }

        private static string GetSourceId(int index)
        {
            return "S" + index.ToString(CultureInfo.InvariantCulture);
        }

        private static string AppendSourceLinks(
            string answer,
            IList<MoodleEmbeddingSearchResult> searchResults,
            AiSearchOptions aiSearchOptions)
        {
            if (searchResults == null || searchResults.Count == 0)
                return answer;

            MoodleEmbeddingSearchResult primaryResult = GetPrimarySourceResults(searchResults).FirstOrDefault();
            if (primaryResult == null)
                return answer;

            List<SourceLinkInfo> sourceLinks = BuildArticleSourceLinks(primaryResult);
            if (sourceLinks.Count == 0)
                return answer;

            StringBuilder builder = new StringBuilder(RemoveVisibleSourceUrlMentions(answer, sourceLinks).TrimEnd());
            builder.Append("<br><br>");
            for (int i = 0; i < sourceLinks.Count; i++)
            {
                if (i > 0)
                    builder.Append(" ");

                SourceLinkInfo sourceLink = sourceLinks[i];
                builder.Append("<a data-source=\"true\" data-source-label=\"")
                    .Append(WebUtility.HtmlEncode(GetSourceButtonText(aiSearchOptions, i + 1, sourceLinks.Count, sourceLink.Score)))
                    .Append("\" href=\"")
                    .Append(WebUtility.HtmlEncode(sourceLink.Url))
                    .Append("\"></a>");
            }

            return builder.ToString();
        }

        private static string RemoveVisibleSourceUrlMentions(string answer, IEnumerable<SourceLinkInfo> sourceLinks)
        {
            string cleanAnswer = answer ?? string.Empty;
            if (sourceLinks == null)
                return cleanAnswer;

            HashSet<string> urlVariants = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (SourceLinkInfo sourceLink in sourceLinks)
            {
                if (sourceLink == null)
                    continue;

                foreach (string urlVariant in GetVisibleSourceUrlVariants(sourceLink.Url))
                    urlVariants.Add(urlVariant);
            }

            foreach (string urlVariant in urlVariants.OrderByDescending(value => value.Length))
            {
                if (string.IsNullOrWhiteSpace(urlVariant))
                    continue;

                cleanAnswer = Regex.Replace(
                    cleanAnswer,
                    "(?<![\"'])\\s*(?:[-–—]\\s*)?" + Regex.Escape(urlVariant) + "\\s*(?:[-–—]\\s*)?",
                    " ",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            }

            return CleanSourceUrlRemovalArtifacts(cleanAnswer);
        }

        private static IEnumerable<string> GetVisibleSourceUrlVariants(string sourceUrl)
        {
            string normalizedSourceUrl = NormalizeSourceUrl(sourceUrl);
            if (string.IsNullOrWhiteSpace(normalizedSourceUrl))
                yield break;

            yield return normalizedSourceUrl;
            yield return WebUtility.HtmlEncode(normalizedSourceUrl);

            Uri uri;
            if (!Uri.TryCreate(normalizedSourceUrl, UriKind.Absolute, out uri))
                yield break;

            yield return uri.AbsoluteUri;
            yield return WebUtility.HtmlEncode(uri.AbsoluteUri);

            string unescapedUrl;
            if (TryUnescapeUriString(uri.AbsoluteUri, out unescapedUrl))
            {
                yield return unescapedUrl;
                yield return WebUtility.HtmlEncode(unescapedUrl);
            }
        }

        private static bool TryUnescapeUriString(string value, out string result)
        {
            result = string.Empty;
            if (string.IsNullOrWhiteSpace(value))
                return false;

            try
            {
                result = Uri.UnescapeDataString(value);
                return !string.IsNullOrWhiteSpace(result);
            }
            catch
            {
                return false;
            }
        }

        private static string CleanSourceUrlRemovalArtifacts(string answer)
        {
            string cleanAnswer = answer ?? string.Empty;
            cleanAnswer = Regex.Replace(cleanAnswer, "[\\(\\[]\\s*[\\)\\]]", string.Empty, RegexOptions.CultureInvariant);
            cleanAnswer = Regex.Replace(cleanAnswer, "[ \\t]+(\\r?\\n)", "$1", RegexOptions.CultureInvariant);
            cleanAnswer = Regex.Replace(cleanAnswer, "[ \\t]{2,}", " ", RegexOptions.CultureInvariant);
            cleanAnswer = Regex.Replace(cleanAnswer, "\\s+([,.;:!?])", "$1", RegexOptions.CultureInvariant);
            cleanAnswer = Regex.Replace(cleanAnswer, "([\\(\\[])\\s+", "$1", RegexOptions.CultureInvariant);
            cleanAnswer = Regex.Replace(cleanAnswer, "\\s+([\\)\\]])", "$1", RegexOptions.CultureInvariant);
            cleanAnswer = Regex.Replace(cleanAnswer, "(?<=[.!?:;])(?=\\S)", " ", RegexOptions.CultureInvariant);
            return cleanAnswer.Trim();
        }

        private static List<SourceLinkInfo> BuildArticleSourceLinks(MoodleEmbeddingSearchResult primaryResult)
        {
            List<SourceLinkInfo> sourceLinks = new List<SourceLinkInfo>();
            if (primaryResult == null || primaryResult.Article == null)
                return sourceLinks;

            HashSet<string> sourceKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            AddSourceLink(sourceLinks, sourceKeys, primaryResult.Article.SourceUrl, primaryResult.Score);
            foreach (string htmlSourceUrl in ExtractHtmlSourceUrls(primaryResult.Article.Html, primaryResult.Article.SourceUrl))
                AddSourceLink(sourceLinks, sourceKeys, htmlSourceUrl, double.NaN);

            return sourceLinks;
        }

        private static void AddSourceLink(
            ICollection<SourceLinkInfo> sourceLinks,
            ISet<string> sourceKeys,
            string sourceUrl,
            double score)
        {
            string normalizedSourceUrl = NormalizeSourceUrl(sourceUrl);
            if (string.IsNullOrWhiteSpace(normalizedSourceUrl))
                return;

            Uri uri;
            if (!Uri.TryCreate(normalizedSourceUrl, UriKind.Absolute, out uri) ||
                (!string.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) &&
                 !string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }

            string sourceKey = GetSourceDedupKey(uri.AbsoluteUri);
            if (string.IsNullOrWhiteSpace(sourceKey) || !sourceKeys.Add(sourceKey))
                return;

            sourceLinks.Add(new SourceLinkInfo(uri.AbsoluteUri, score));
        }

        private static IEnumerable<string> ExtractHtmlSourceUrls(string html, string baseSourceUrl)
        {
            if (string.IsNullOrWhiteSpace(html))
                yield break;

            foreach (Match anchorMatch in HtmlAnchorTagRegex.Matches(html))
            {
                Match hrefMatch = HtmlHrefAttributeRegex.Match(anchorMatch.Groups["attrs"].Value ?? string.Empty);
                if (!hrefMatch.Success)
                    continue;

                string resolvedUrl;
                if (TryResolveHtmlSourceUrl(hrefMatch.Groups["href"].Value, baseSourceUrl, out resolvedUrl))
                    yield return resolvedUrl;
            }

            foreach (Match urlMatch in PlainHttpUrlRegex.Matches(html))
            {
                string resolvedUrl;
                if (TryResolveHtmlSourceUrl(urlMatch.Value, baseSourceUrl, out resolvedUrl))
                    yield return resolvedUrl;
            }
        }

        private static bool TryResolveHtmlSourceUrl(string sourceUrl, string baseSourceUrl, out string resolvedUrl)
        {
            resolvedUrl = string.Empty;
            string normalizedSourceUrl = NormalizeSourceUrl(sourceUrl).TrimEnd('.', ',', ';', ')', ']');
            if (string.IsNullOrWhiteSpace(normalizedSourceUrl) ||
                normalizedSourceUrl.StartsWith("#", StringComparison.Ordinal) ||
                normalizedSourceUrl.StartsWith("javascript:", StringComparison.OrdinalIgnoreCase) ||
                normalizedSourceUrl.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase) ||
                normalizedSourceUrl.StartsWith("tel:", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            Uri uri;
            if (!Uri.TryCreate(normalizedSourceUrl, UriKind.Absolute, out uri))
            {
                Uri baseUri;
                if (!Uri.TryCreate(NormalizeSourceUrl(baseSourceUrl), UriKind.Absolute, out baseUri) ||
                    !Uri.TryCreate(baseUri, normalizedSourceUrl, out uri))
                {
                    return false;
                }
            }

            if (!string.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            resolvedUrl = uri.AbsoluteUri;
            return true;
        }

        private static string NormalizeSourceUrl(string sourceUrl)
        {
            return WebUtility.HtmlDecode(sourceUrl ?? string.Empty).Trim();
        }

        private static IList<MoodleEmbeddingSearchResult> GetUniqueSourceResults(
            IList<MoodleEmbeddingSearchResult> searchResults,
            int maxCount)
        {
            List<MoodleEmbeddingSearchResult> uniqueResults = new List<MoodleEmbeddingSearchResult>();
            if (searchResults == null || searchResults.Count == 0 || maxCount <= 0)
                return uniqueResults;

            HashSet<string> sourceKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (MoodleEmbeddingSearchResult result in searchResults)
            {
                if (result == null || result.Article == null || string.IsNullOrWhiteSpace(result.Article.SourceUrl))
                    continue;

                string sourceKey = GetSourceDedupKey(result.Article.SourceUrl);
                if (string.IsNullOrWhiteSpace(sourceKey) || !sourceKeys.Add(sourceKey))
                    continue;

                uniqueResults.Add(result);
                if (uniqueResults.Count >= maxCount)
                    break;
            }

            return uniqueResults;
        }

        private static string GetSourceDedupKey(string sourceUrl)
        {
            string decodedSourceUrl = NormalizeSourceUrl(sourceUrl);
            if (string.IsNullOrWhiteSpace(decodedSourceUrl))
                return string.Empty;

            Uri uri;
            if (!Uri.TryCreate(decodedSourceUrl, UriKind.Absolute, out uri))
                return decodedSourceUrl;

            return uri.GetLeftPart(UriPartial.Path).TrimEnd('/') + "?" + uri.Query.TrimStart('?');
        }

        private static string GetSourceButtonText(AiSearchOptions aiSearchOptions, int index, int totalCount, double score)
        {
            string sourceButtonText = aiSearchOptions == null || string.IsNullOrWhiteSpace(aiSearchOptions.SourceButtonText)
                ? "Источник"
                : aiSearchOptions.SourceButtonText.Trim();
            if (totalCount > 1)
                sourceButtonText += " " + index;

            if (aiSearchOptions != null && aiSearchOptions.DebugRetrieval && !double.IsNaN(score))
            {
                sourceButtonText += " (" + score.ToString("0.00", CultureInfo.InvariantCulture) + ")";
            }

            return sourceButtonText;
        }

        private sealed class UsedSourceMarker
        {
            public UsedSourceMarker(string cleanAnswer, IList<string> sourceIds, bool found)
            {
                CleanAnswer = cleanAnswer;
                SourceIds = sourceIds;
                Found = found;
            }

            public string CleanAnswer { get; private set; }

            public IList<string> SourceIds { get; private set; }

            public bool Found { get; private set; }
        }

        private sealed class SourceLinkInfo
        {
            public SourceLinkInfo(string url, double score)
            {
                Url = url;
                Score = score;
            }

            public string Url { get; private set; }

            public double Score { get; private set; }
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
        private sealed class GigaChatEmbeddingRequest
        {
            [DataMember(Name = "model")]
            public string Model { get; set; }

            [DataMember(Name = "input")]
            public string Input { get; set; }
        }

        [DataContract]
        private sealed class GigaChatEmbeddingResponse
        {
            [DataMember(Name = "data")]
            public List<GigaChatEmbeddingData> Data { get; set; }
        }

        [DataContract]
        private sealed class GigaChatEmbeddingData
        {
            [DataMember(Name = "embedding")]
            public List<double> Embedding { get; set; }
        }

        [DataContract]
        private sealed class GigaChatChoice
        {
            [DataMember(Name = "message")]
            public GigaChatMessage Message { get; set; }
        }
    }
}