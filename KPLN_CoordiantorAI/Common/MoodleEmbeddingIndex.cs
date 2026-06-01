using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;

namespace KPLN_CoordiantorAI.Common
{
    public sealed class MoodleEmbeddingIndex
    {
        private const string MetadataFileName = "moodle_meta.json";
        private const string EmbeddingsJsonFileName = "moodle_embeddings.json";
        private const string EmbeddingsNpyFileName = "moodle_embeddings.npy";

        private static readonly Regex NpyShapeRegex = new Regex(
            "\\(\\s*(?<rows>\\d+)\\s*,\\s*(?<cols>\\d+)\\s*,?\\s*\\)",
            RegexOptions.CultureInvariant);

        private readonly IList<MoodleEmbeddingArticle> _articles;

        private MoodleEmbeddingIndex(IList<MoodleEmbeddingArticle> articles)
        {
            _articles = articles;
        }

        public static MoodleEmbeddingIndex Load(string embeddingsFolderPath)
        {
            return Load(embeddingsFolderPath, AiSearchOptions.DefaultMoodleUrlTemplate);
        }

        public static MoodleEmbeddingIndex Load(string embeddingsFolderPath, string moodleUrlTemplate)
        {
            string folderPath = (embeddingsFolderPath ?? string.Empty).Trim().Trim('"');
            if (string.IsNullOrWhiteSpace(folderPath))
                return null;

            if (!Directory.Exists(folderPath))
                throw new InvalidOperationException("Embeddings folder was not found: " + folderPath);

            string metadataPath = Path.Combine(folderPath, MetadataFileName);
            string embeddingsNpyPath = Path.Combine(folderPath, EmbeddingsNpyFileName);
            string embeddingsJsonPath = Path.Combine(folderPath, EmbeddingsJsonFileName);
            if (!File.Exists(metadataPath) || (!File.Exists(embeddingsNpyPath) && !File.Exists(embeddingsJsonPath)))
            {
                throw new InvalidOperationException(
                    "Embeddings folder must contain " + MetadataFileName + " and " + EmbeddingsNpyFileName + " or " + EmbeddingsJsonFileName + ".");
            }

            List<MoodleEmbeddingMetadata> metadata = Deserialize<List<MoodleEmbeddingMetadata>>(File.ReadAllText(metadataPath, Encoding.UTF8));
            List<List<double>> embeddings = File.Exists(embeddingsNpyPath)
                ? LoadNpyFloat32Matrix(embeddingsNpyPath)
                : Deserialize<List<List<double>>>(File.ReadAllText(embeddingsJsonPath, Encoding.UTF8));
            if (metadata == null || embeddings == null || metadata.Count == 0 || embeddings.Count == 0)
                throw new InvalidOperationException("Moodle embeddings index is empty.");

            if (metadata.Count != embeddings.Count)
                throw new InvalidOperationException("Moodle meta count does not match embedding vector count.");

            List<MoodleEmbeddingArticle> articles = new List<MoodleEmbeddingArticle>();
            for (int i = 0; i < metadata.Count; i++)
            {
                MoodleEmbeddingMetadata item = metadata[i];
                List<double> vector = embeddings[i];
                if (item == null || vector == null || vector.Count == 0)
                    continue;

                string sourceUrl = item.GetSourceUrl(moodleUrlTemplate);
                if (string.IsNullOrWhiteSpace(sourceUrl))
                    continue;

                articles.Add(new MoodleEmbeddingArticle(
                    item.ArticleId,
                    item.Title,
                    item.Text,
                    item.GetHeadings(),
                    item.GetSections(),
                    item.GetHtml(folderPath),
                    sourceUrl,
                    NormalizeVector(vector)));
            }

            if (articles.Count == 0)
            {
                throw new InvalidOperationException(
                    "Moodle meta does not contain articles with source links. Rebuild embeddings with the Python script.");
            }

            return new MoodleEmbeddingIndex(articles);
        }

        public static IList<MoodleArticleInfo> LoadArticleInfos(string embeddingsFolderPath)
        {
            return LoadArticleInfos(embeddingsFolderPath, AiSearchOptions.DefaultMoodleUrlTemplate);
        }

        public static IList<MoodleArticleInfo> LoadArticleInfos(string embeddingsFolderPath, string moodleUrlTemplate)
        {
            List<MoodleArticleInfo> articles = new List<MoodleArticleInfo>();
            string folderPath = (embeddingsFolderPath ?? string.Empty).Trim().Trim('"');
            if (string.IsNullOrWhiteSpace(folderPath))
                return articles;

            if (!Directory.Exists(folderPath))
                throw new InvalidOperationException("Embeddings folder was not found: " + folderPath);

            string metadataPath = Path.Combine(folderPath, MetadataFileName);
            if (!File.Exists(metadataPath))
                throw new InvalidOperationException("Embeddings folder must contain " + MetadataFileName + ".");

            List<MoodleEmbeddingMetadata> metadata = Deserialize<List<MoodleEmbeddingMetadata>>(File.ReadAllText(metadataPath, Encoding.UTF8));
            if (metadata == null)
                return articles;

            HashSet<string> knownKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (MoodleEmbeddingMetadata item in metadata)
            {
                if (item == null)
                    continue;

                string sourceUrl = item.GetSourceUrl(moodleUrlTemplate);
                string key = !string.IsNullOrWhiteSpace(item.ArticleId) ? item.ArticleId.Trim() : sourceUrl;
                if (string.IsNullOrWhiteSpace(key) || !knownKeys.Add(key))
                    continue;

                string title = string.IsNullOrWhiteSpace(item.Title) ? item.HtmlFile : item.Title;
                articles.Add(new MoodleArticleInfo(
                    item.ArticleId,
                    string.IsNullOrWhiteSpace(title) ? key : title.Trim(),
                    sourceUrl,
                    item.HtmlFile));
            }

            return articles;
        }

        public MoodleEmbeddingArticle FindBest(IList<double> queryVector)
        {
            IList<MoodleEmbeddingSearchResult> results = FindTop(queryVector, 1, double.MinValue, string.Empty);
            return results.Count == 0 ? null : results[0].Article;
        }

        public IList<MoodleEmbeddingSearchResult> FindTop(IList<double> queryVector, int topK, double minSimilarity)
        {
            return FindTop(queryVector, topK, minSimilarity, string.Empty);
        }

        public IList<MoodleEmbeddingSearchResult> FindTop(
            IList<double> queryVector,
            int topK,
            double minSimilarity,
            string queryText)
        {
            List<MoodleEmbeddingSearchResult> results = new List<MoodleEmbeddingSearchResult>();
            if (queryVector == null || queryVector.Count == 0 || _articles == null || _articles.Count == 0)
                return results;

            if (topK < 1)
                topK = 1;

            IList<double> normalizedQueryVector = NormalizeVector(queryVector);
            foreach (MoodleEmbeddingArticle article in _articles)
            {
                double score = DotProduct(normalizedQueryVector, article.Vector);
                if (score < minSimilarity)
                    continue;

                results.Add(new MoodleEmbeddingSearchResult(article, score, article.GetRelevantSections(queryText, 3)));
            }

            results.Sort((first, second) => second.Score.CompareTo(first.Score));
            if (results.Count > topK)
                results.RemoveRange(topK, results.Count - topK);

            return results;
        }

        public IList<MoodleEmbeddingSearchResult> FindByArticleIds(IList<string> articleIds, string queryText)
        {
            List<MoodleEmbeddingSearchResult> results = new List<MoodleEmbeddingSearchResult>();
            if (articleIds == null || articleIds.Count == 0 || _articles == null || _articles.Count == 0)
                return results;

            HashSet<string> usedArticleIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string articleId in articleIds)
            {
                string normalizedArticleId = (articleId ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(normalizedArticleId) || !usedArticleIds.Add(normalizedArticleId))
                    continue;

                foreach (MoodleEmbeddingArticle article in _articles)
                {
                    if (article == null || !string.Equals(article.ArticleId, normalizedArticleId, StringComparison.OrdinalIgnoreCase))
                        continue;

                    results.Add(new MoodleEmbeddingSearchResult(article, 1.0d, article.GetRelevantSections(queryText, 3)));
                    break;
                }
            }

            return results;
        }

        private static double DotProduct(IList<double> first, IList<double> second)
        {
            if (first == null || second == null || first.Count != second.Count || first.Count == 0)
                return double.MinValue;

            double dot = 0;
            for (int i = 0; i < first.Count; i++)
                dot += first[i] * second[i];

            return dot;
        }

        private static IList<double> NormalizeVector(IList<double> vector)
        {
            double norm = 0;
            for (int i = 0; i < vector.Count; i++)
                norm += vector[i] * vector[i];

            norm = Math.Sqrt(norm);
            if (norm <= 0)
                return vector;

            double[] normalized = new double[vector.Count];
            for (int i = 0; i < vector.Count; i++)
                normalized[i] = vector[i] / norm;

            return normalized;
        }

        private static List<List<double>> LoadNpyFloat32Matrix(string filePath)
        {
            using (BinaryReader reader = new BinaryReader(File.OpenRead(filePath)))
            {
                byte[] magic = reader.ReadBytes(6);
                if (magic.Length != 6 ||
                    magic[0] != 0x93 ||
                    magic[1] != (byte)'N' ||
                    magic[2] != (byte)'U' ||
                    magic[3] != (byte)'M' ||
                    magic[4] != (byte)'P' ||
                    magic[5] != (byte)'Y')
                {
                    throw new InvalidOperationException("Invalid npy file header: " + filePath);
                }

                byte majorVersion = reader.ReadByte();
                reader.ReadByte();
                int headerLength = majorVersion == 1
                    ? reader.ReadUInt16()
                    : (int)reader.ReadUInt32();
                string header = Encoding.ASCII.GetString(reader.ReadBytes(headerLength));

                if (header.IndexOf("fortran_order", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    header.IndexOf("True", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    throw new InvalidOperationException("Fortran-order npy embeddings are not supported.");
                }

                bool isLittleEndian = header.IndexOf("'<f4'", StringComparison.Ordinal) >= 0 ||
                    header.IndexOf("\"<f4\"", StringComparison.Ordinal) >= 0 ||
                    header.IndexOf("'|f4'", StringComparison.Ordinal) >= 0 ||
                    header.IndexOf("\"|f4\"", StringComparison.Ordinal) >= 0;
                bool isBigEndian = header.IndexOf("'>f4'", StringComparison.Ordinal) >= 0 ||
                    header.IndexOf("\">f4\"", StringComparison.Ordinal) >= 0;
                if (!isLittleEndian && !isBigEndian)
                    throw new InvalidOperationException("Only float32 npy embeddings are supported.");

                Match shapeMatch = NpyShapeRegex.Match(header);
                if (!shapeMatch.Success)
                    throw new InvalidOperationException("Cannot read npy matrix shape.");

                int rowCount = int.Parse(shapeMatch.Groups["rows"].Value);
                int columnCount = int.Parse(shapeMatch.Groups["cols"].Value);
                List<List<double>> rows = new List<List<double>>(rowCount);
                for (int row = 0; row < rowCount; row++)
                {
                    List<double> values = new List<double>(columnCount);
                    for (int column = 0; column < columnCount; column++)
                    {
                        byte[] bytes = reader.ReadBytes(4);
                        if (bytes.Length != 4)
                            throw new InvalidOperationException("Unexpected end of npy embeddings file.");

                        if (isBigEndian == BitConverter.IsLittleEndian)
                            Array.Reverse(bytes);

                        values.Add(BitConverter.ToSingle(bytes, 0));
                    }

                    rows.Add(values);
                }

                return rows;
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
        private sealed class MoodleEmbeddingMetadata
        {
            [DataMember(Name = "article_id")]
            public string ArticleId { get; set; }

            [DataMember(Name = "html")]
            public string Html { get; set; }

            [DataMember(Name = "source_url")]
            public string SourceUrl { get; set; }

            [DataMember(Name = "html_file")]
            public string HtmlFile { get; set; }

            [DataMember(Name = "book_id")]
            public int BookId { get; set; }

            [DataMember(Name = "chapter_id")]
            public int ChapterId { get; set; }

            [DataMember(Name = "title")]
            public string Title { get; set; }

            [DataMember(Name = "text")]
            public string Text { get; set; }

            [DataMember(Name = "headings")]
            public List<MoodleEmbeddingMetadataHeading> Headings { get; set; }

            [DataMember(Name = "sections")]
            public List<MoodleEmbeddingMetadataSection> Sections { get; set; }

            public IList<MoodleEmbeddingHeading> GetHeadings()
            {
                List<MoodleEmbeddingHeading> headings = new List<MoodleEmbeddingHeading>();
                if (Headings == null)
                    return headings;

                foreach (MoodleEmbeddingMetadataHeading heading in Headings)
                {
                    if (heading == null || string.IsNullOrWhiteSpace(heading.Text))
                        continue;

                    headings.Add(new MoodleEmbeddingHeading(heading.Level, heading.Text));
                }

                return headings;
            }

            public IList<MoodleEmbeddingSection> GetSections()
            {
                List<MoodleEmbeddingSection> sections = new List<MoodleEmbeddingSection>();
                if (Sections == null)
                    return sections;

                foreach (MoodleEmbeddingMetadataSection section in Sections)
                {
                    if (section == null || string.IsNullOrWhiteSpace(section.Text))
                        continue;

                    sections.Add(new MoodleEmbeddingSection(section.HeadingPath, section.Text));
                }

                return sections;
            }

            public string GetHtml(string embeddingsFolderPath)
            {
                if (!string.IsNullOrWhiteSpace(Html))
                    return Html;

                if (string.IsNullOrWhiteSpace(HtmlFile))
                    return string.Empty;

                string directHtmlPath = Path.Combine(embeddingsFolderPath, "html", HtmlFile);
                if (File.Exists(directHtmlPath))
                    return File.ReadAllText(directHtmlPath, Encoding.UTF8);

                string outputFolder = Path.GetDirectoryName(embeddingsFolderPath);
                string siblingHtmlPath = Path.Combine(outputFolder ?? embeddingsFolderPath, "html", HtmlFile);
                return File.Exists(siblingHtmlPath) ? File.ReadAllText(siblingHtmlPath, Encoding.UTF8) : string.Empty;
            }

            public string GetSourceUrl(string moodleUrlTemplate)
            {
                if (!string.IsNullOrWhiteSpace(SourceUrl))
                {
                    string sourceUrl = SourceUrl.Trim();
                    string templatedSourceUrl = TryBuildSourceUrlFromTemplate(moodleUrlTemplate, sourceUrl);
                    return string.IsNullOrWhiteSpace(templatedSourceUrl) ? sourceUrl : templatedSourceUrl;
                }

                if (BookId <= 0 || ChapterId <= 0)
                    return string.Empty;

                return BuildSourceUrl(moodleUrlTemplate, BookId, ChapterId);
            }

            private static string BuildSourceUrl(string moodleUrlTemplate, int bookId, int chapterId)
            {
                string template = string.IsNullOrWhiteSpace(moodleUrlTemplate)
                    ? AiSearchOptions.DefaultMoodleUrlTemplate
                    : moodleUrlTemplate.Trim();

                return template
                    .Replace("{book_id}", bookId.ToString())
                    .Replace("{chapter_id}", chapterId.ToString());
            }

            private static string TryBuildSourceUrlFromTemplate(string moodleUrlTemplate, string sourceUrl)
            {
                int bookId;
                int chapterId;
                if (!TryGetQueryInt(sourceUrl, "id", out bookId) ||
                    !TryGetQueryInt(sourceUrl, "chapterid", out chapterId))
                {
                    return string.Empty;
                }

                return BuildSourceUrl(moodleUrlTemplate, bookId, chapterId);
            }

            private static bool TryGetQueryInt(string sourceUrl, string parameterName, out int value)
            {
                value = 0;
                string query = GetQuery(sourceUrl);
                if (string.IsNullOrWhiteSpace(query))
                    return false;

                string[] parts = query.Split('&');
                foreach (string part in parts)
                {
                    if (string.IsNullOrWhiteSpace(part))
                        continue;

                    int separatorIndex = part.IndexOf('=');
                    string key = separatorIndex < 0 ? part : part.Substring(0, separatorIndex);
                    string rawValue = separatorIndex < 0 ? string.Empty : part.Substring(separatorIndex + 1);
                    if (!string.Equals(Uri.UnescapeDataString(key), parameterName, StringComparison.OrdinalIgnoreCase))
                        continue;

                    return int.TryParse(Uri.UnescapeDataString(rawValue), out value);
                }

                return false;
            }

            private static string GetQuery(string sourceUrl)
            {
                if (string.IsNullOrWhiteSpace(sourceUrl))
                    return string.Empty;

                sourceUrl = sourceUrl.Replace("&amp;", "&").Replace("&AMP;", "&");

                Uri uri;
                if (Uri.TryCreate(sourceUrl, UriKind.Absolute, out uri))
                    return uri.Query.TrimStart('?');

                int queryIndex = sourceUrl.IndexOf('?');
                return queryIndex < 0 || queryIndex >= sourceUrl.Length - 1
                    ? string.Empty
                    : sourceUrl.Substring(queryIndex + 1);
            }
        }

        [DataContract]
        private sealed class MoodleEmbeddingMetadataHeading
        {
            [DataMember(Name = "level")]
            public int Level { get; set; }

            [DataMember(Name = "text")]
            public string Text { get; set; }
        }

        [DataContract]
        private sealed class MoodleEmbeddingMetadataSection
        {
            [DataMember(Name = "heading_path")]
            public List<string> HeadingPath { get; set; }

            [DataMember(Name = "text")]
            public string Text { get; set; }
        }
    }

    public sealed class MoodleEmbeddingArticle
    {
        public MoodleEmbeddingArticle(
            string articleId,
            string title,
            string text,
            IList<MoodleEmbeddingHeading> headings,
            IList<MoodleEmbeddingSection> sections,
            string html,
            string sourceUrl,
            IList<double> vector)
        {
            ArticleId = articleId;
            Title = title;
            Text = text;
            Headings = headings ?? new List<MoodleEmbeddingHeading>();
            Sections = sections ?? new List<MoodleEmbeddingSection>();
            Html = html;
            SourceUrl = sourceUrl;
            Vector = vector;
        }

        public string ArticleId { get; private set; }

        public string Title { get; private set; }

        public string Text { get; private set; }

        public IList<MoodleEmbeddingHeading> Headings { get; private set; }

        public IList<MoodleEmbeddingSection> Sections { get; private set; }

        public string Html { get; private set; }

        public string SourceUrl { get; private set; }

        public IList<double> Vector { get; private set; }

        public string GetHeadingsText()
        {
            if (Headings == null || Headings.Count == 0)
                return string.Empty;

            StringBuilder builder = new StringBuilder();
            foreach (MoodleEmbeddingHeading heading in Headings)
            {
                if (heading == null || string.IsNullOrWhiteSpace(heading.Text))
                    continue;

                if (builder.Length > 0)
                    builder.Append(" > ");

                builder.Append("H").Append(heading.Level).Append(": ").Append(heading.Text.Trim());
            }

            return builder.ToString();
        }

        public IList<MoodleEmbeddingSection> GetRelevantSections(string queryText, int maxCount)
        {
            List<MoodleEmbeddingSection> relevantSections = new List<MoodleEmbeddingSection>();
            if (Sections == null || Sections.Count == 0 || maxCount <= 0)
                return relevantSections;

            HashSet<string> queryTokens = Tokenize(queryText);
            List<SectionScore> scores = new List<SectionScore>();
            foreach (MoodleEmbeddingSection section in Sections)
            {
                if (section == null || string.IsNullOrWhiteSpace(section.Text))
                    continue;

                scores.Add(new SectionScore(section, GetSectionScore(section, queryTokens)));
            }

            scores.Sort((first, second) => second.Score.CompareTo(first.Score));
            foreach (SectionScore score in scores)
            {
                if (relevantSections.Count >= maxCount)
                    break;

                if (queryTokens.Count > 0 && score.Score <= 0)
                    continue;

                relevantSections.Add(score.Section);
            }

            if (relevantSections.Count > 0)
                return relevantSections;

            foreach (MoodleEmbeddingSection section in Sections)
            {
                if (section == null || string.IsNullOrWhiteSpace(section.Text))
                    continue;

                relevantSections.Add(section);
                if (relevantSections.Count >= maxCount)
                    break;
            }

            return relevantSections;
        }

        private static double GetSectionScore(MoodleEmbeddingSection section, HashSet<string> queryTokens)
        {
            if (queryTokens == null || queryTokens.Count == 0)
                return 0;

            string sectionText = NormalizeText(section.GetHeadingPathText() + " " + section.Text);
            double score = 0;
            foreach (string token in queryTokens)
            {
                if (sectionText.IndexOf(token, StringComparison.Ordinal) >= 0)
                    score += 1;
            }

            return score;
        }

        private static HashSet<string> Tokenize(string value)
        {
            HashSet<string> tokens = new HashSet<string>(StringComparer.Ordinal);
            StringBuilder token = new StringBuilder();
            string normalized = NormalizeText(value);
            foreach (char currentChar in normalized)
            {
                if (char.IsLetterOrDigit(currentChar))
                {
                    token.Append(currentChar);
                    continue;
                }

                AddToken(tokens, token);
            }

            AddToken(tokens, token);
            return tokens;
        }

        private static void AddToken(ISet<string> tokens, StringBuilder token)
        {
            if (token.Length >= 3)
                tokens.Add(token.ToString());

            token.Length = 0;
        }

        private static string NormalizeText(string value)
        {
            return (value ?? string.Empty).Trim().ToLowerInvariant().Replace('ё', 'е');
        }

        private sealed class SectionScore
        {
            public SectionScore(MoodleEmbeddingSection section, double score)
            {
                Section = section;
                Score = score;
            }

            public MoodleEmbeddingSection Section { get; private set; }

            public double Score { get; private set; }
        }
    }

    public sealed class MoodleEmbeddingHeading
    {
        public MoodleEmbeddingHeading(int level, string text)
        {
            Level = level;
            Text = text;
        }

        public int Level { get; private set; }

        public string Text { get; private set; }
    }

    public sealed class MoodleEmbeddingSection
    {
        public MoodleEmbeddingSection(IList<string> headingPath, string text)
        {
            HeadingPath = headingPath ?? new List<string>();
            Text = text;
        }

        public IList<string> HeadingPath { get; private set; }

        public string Text { get; private set; }

        public string GetHeadingPathText()
        {
            if (HeadingPath == null || HeadingPath.Count == 0)
                return string.Empty;

            StringBuilder builder = new StringBuilder();
            foreach (string heading in HeadingPath)
            {
                if (string.IsNullOrWhiteSpace(heading))
                    continue;

                if (builder.Length > 0)
                    builder.Append(" > ");

                builder.Append(heading.Trim());
            }

            return builder.ToString();
        }
    }

    public sealed class MoodleArticleInfo
    {
        public MoodleArticleInfo(string articleId, string title, string sourceUrl, string htmlFile)
        {
            ArticleId = articleId;
            Title = title;
            SourceUrl = sourceUrl;
            HtmlFile = htmlFile;
        }

        public string ArticleId { get; private set; }

        public string Title { get; private set; }

        public string SourceUrl { get; private set; }

        public string HtmlFile { get; private set; }
    }

    public sealed class MoodleEmbeddingSearchResult
    {
        public MoodleEmbeddingSearchResult(
            MoodleEmbeddingArticle article,
            double score,
            IList<MoodleEmbeddingSection> relevantSections)
        {
            Article = article;
            Score = score;
            RelevantSections = relevantSections ?? new List<MoodleEmbeddingSection>();
        }

        public MoodleEmbeddingArticle Article { get; private set; }

        public double Score { get; private set; }

        public IList<MoodleEmbeddingSection> RelevantSections { get; private set; }
    }
}