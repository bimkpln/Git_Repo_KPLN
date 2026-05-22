using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace KPLN_CoordiantorAI.Common
{
    public sealed class MoodleEmbeddingIndex
    {
        private const string MetadataFileName = "moodle_meta.json";
        private const string EmbeddingsFileName = "moodle_embeddings.json";

        private readonly IList<MoodleEmbeddingArticle> _articles;

        private MoodleEmbeddingIndex(IList<MoodleEmbeddingArticle> articles)
        {
            _articles = articles;
        }

        public static MoodleEmbeddingIndex Load(string embeddingsFolderPath)
        {
            string folderPath = (embeddingsFolderPath ?? string.Empty).Trim().Trim('"');
            if (string.IsNullOrWhiteSpace(folderPath))
                return null;

            if (!Directory.Exists(folderPath))
                throw new InvalidOperationException("Папка с эмбеддингами не найдена: " + folderPath);

            string metadataPath = Path.Combine(folderPath, MetadataFileName);
            string embeddingsPath = Path.Combine(folderPath, EmbeddingsFileName);
            if (!File.Exists(metadataPath) || !File.Exists(embeddingsPath))
            {
                throw new InvalidOperationException(
                    "В папке с эмбеддингами нужны файлы " + MetadataFileName + " и " + EmbeddingsFileName + ".");
            }

            List<MoodleEmbeddingMetadata> metadata = Deserialize<List<MoodleEmbeddingMetadata>>(File.ReadAllText(metadataPath));
            List<List<double>> embeddings = Deserialize<List<List<double>>>(File.ReadAllText(embeddingsPath));
            if (metadata == null || embeddings == null || metadata.Count == 0 || embeddings.Count == 0)
                throw new InvalidOperationException("Индекс эмбеддингов Moodle пуст.");

            if (metadata.Count != embeddings.Count)
                throw new InvalidOperationException("Количество Moodle meta и embedding vectors не совпадает.");

            List<MoodleEmbeddingArticle> articles = new List<MoodleEmbeddingArticle>();
            for (int i = 0; i < metadata.Count; i++)
            {
                MoodleEmbeddingMetadata item = metadata[i];
                List<double> vector = embeddings[i];
                if (item == null || vector == null || vector.Count == 0)
                    continue;

                string html = item.GetHtml(folderPath);
                string sourceUrl = item.GetSourceUrl();
                if (string.IsNullOrWhiteSpace(html) || string.IsNullOrWhiteSpace(sourceUrl))
                    continue;

                articles.Add(new MoodleEmbeddingArticle(html, sourceUrl, vector.ToArray()));
            }

            if (articles.Count == 0)
            {
                throw new InvalidOperationException(
                    "В Moodle meta нет HTML статьи и ссылки. Пересоберите эмбеддинги Python-скриптом.");
            }

            return new MoodleEmbeddingIndex(articles);
        }

        public MoodleEmbeddingArticle FindBest(IList<double> queryVector)
        {
            if (queryVector == null || queryVector.Count == 0 || _articles == null || _articles.Count == 0)
                return null;

            MoodleEmbeddingArticle bestArticle = null;
            double bestScore = double.MinValue;
            foreach (MoodleEmbeddingArticle article in _articles)
            {
                double score = CosineSimilarity(queryVector, article.Vector);
                if (score <= bestScore)
                    continue;

                bestScore = score;
                bestArticle = article;
            }

            return bestArticle;
        }

        private static double CosineSimilarity(IList<double> first, IList<double> second)
        {
            if (first == null || second == null || first.Count != second.Count || first.Count == 0)
                return double.MinValue;

            double dot = 0;
            double firstNorm = 0;
            double secondNorm = 0;
            for (int i = 0; i < first.Count; i++)
            {
                double firstValue = first[i];
                double secondValue = second[i];
                dot += firstValue * secondValue;
                firstNorm += firstValue * firstValue;
                secondNorm += secondValue * secondValue;
            }

            if (firstNorm <= 0 || secondNorm <= 0)
                return double.MinValue;

            return dot / (Math.Sqrt(firstNorm) * Math.Sqrt(secondNorm));
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

            public string GetHtml(string embeddingsFolderPath)
            {
                if (!string.IsNullOrWhiteSpace(Html))
                    return Html;

                if (string.IsNullOrWhiteSpace(HtmlFile))
                    return string.Empty;

                string outputFolder = Path.GetDirectoryName(embeddingsFolderPath);
                string htmlPath = Path.Combine(outputFolder ?? embeddingsFolderPath, "html", HtmlFile);
                return File.Exists(htmlPath) ? File.ReadAllText(htmlPath) : string.Empty;
            }

            public string GetSourceUrl()
            {
                if (!string.IsNullOrWhiteSpace(SourceUrl))
                    return SourceUrl.Trim();

                if (BookId <= 0 || ChapterId <= 0)
                    return string.Empty;

                return string.Format(
                    "http://moodle/mod/book/view.php?id={0}&chapterid={1}",
                    BookId,
                    ChapterId);
            }
        }
    }

    public sealed class MoodleEmbeddingArticle
    {
        public MoodleEmbeddingArticle(string html, string sourceUrl, IList<double> vector)
        {
            Html = html;
            SourceUrl = sourceUrl;
            Vector = vector;
        }

        public string Html { get; private set; }

        public string SourceUrl { get; private set; }

        public IList<double> Vector { get; private set; }
    }
}