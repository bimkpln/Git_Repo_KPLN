using System;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace KPLN_CoordiantorAI.Common
{
    public sealed class AiSearchOptions
    {
        public const string DefaultMoodleUrlTemplate = "http://moodle/mod/book/view.php?id={book_id}&chapterid={chapter_id}";
        private const string DefaultSourceButtonText = "\u0418\u0441\u0442\u043E\u0447\u043D\u0438\u043A";

        public AiSearchOptions()
        {
            EnableEmbeddingSearch = true;
            TopK = 3;
            MinSimilarity = 0.0;
            MaxContextChars = 16000;
            AllowAnswerWithoutContext = true;
            ChatModel = "GigaChat";
            EmbeddingModel = "Embeddings";
            SourceButtonText = DefaultSourceButtonText;
            MoodleUrlTemplate = DefaultMoodleUrlTemplate;
            DebugRetrieval = false;
        }

        public bool EnableEmbeddingSearch { get; private set; }

        public int TopK { get; private set; }

        public double MinSimilarity { get; private set; }

        public int MaxContextChars { get; private set; }

        public bool AllowAnswerWithoutContext { get; private set; }

        public string ChatModel { get; private set; }

        public string EmbeddingModel { get; private set; }

        public string SourceButtonText { get; private set; }

        public string MoodleUrlTemplate { get; private set; }

        public bool DebugRetrieval { get; private set; }

        public static string DefaultJson
        {
            get
            {
                return "{\r\n" +
                    "  \"enableEmbeddingSearch\": true,\r\n" +
                    "  \"topK\": 3,\r\n" +
                    "  \"minSimilarity\": 0.0,\r\n" +
                    "  \"maxContextChars\": 16000,\r\n" +
                    "  \"allowAnswerWithoutContext\": true,\r\n" +
                    "  \"chatModel\": \"GigaChat\",\r\n" +
                    "  \"embeddingModel\": \"Embeddings\",\r\n" +
                    "  \"sourceButtonText\": \"" + DefaultSourceButtonText + "\",\r\n" +
                    "  \"moodleUrlTemplate\": \"" + DefaultMoodleUrlTemplate + "\",\r\n" +
                    "  \"debugRetrieval\": false\r\n" +
                    "}";
            }
        }

        public static AiSearchOptions FromJson(string json)
        {
            AiSearchOptions options = new AiSearchOptions();
            if (string.IsNullOrWhiteSpace(json))
                return options;

            AiSearchOptionsDto dto;
            try
            {
                DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(AiSearchOptionsDto));
                using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                {
                    dto = serializer.ReadObject(stream) as AiSearchOptionsDto;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("AI search settings JSON is invalid.", ex);
            }

            if (dto == null)
                return options;

            if (dto.EnableEmbeddingSearch.HasValue)
                options.EnableEmbeddingSearch = dto.EnableEmbeddingSearch.Value;
            if (dto.TopK.HasValue)
                options.TopK = dto.TopK.Value;
            if (dto.MinSimilarity.HasValue)
                options.MinSimilarity = dto.MinSimilarity.Value;
            if (dto.MaxContextChars.HasValue)
                options.MaxContextChars = dto.MaxContextChars.Value;
            if (dto.AllowAnswerWithoutContext.HasValue)
                options.AllowAnswerWithoutContext = dto.AllowAnswerWithoutContext.Value;
            if (!string.IsNullOrWhiteSpace(dto.ChatModel))
                options.ChatModel = dto.ChatModel.Trim();
            if (!string.IsNullOrWhiteSpace(dto.EmbeddingModel))
                options.EmbeddingModel = dto.EmbeddingModel.Trim();
            if (!string.IsNullOrWhiteSpace(dto.SourceButtonText))
                options.SourceButtonText = dto.SourceButtonText.Trim();
            if (!string.IsNullOrWhiteSpace(dto.MoodleUrlTemplate))
                options.MoodleUrlTemplate = dto.MoodleUrlTemplate.Trim();
            if (dto.DebugRetrieval.HasValue)
                options.DebugRetrieval = dto.DebugRetrieval.Value;

            options.Normalize();
            return options;
        }

        private void Normalize()
        {
            if (TopK < 1)
                TopK = 1;
            if (TopK > 10)
                TopK = 10;

            if (double.IsNaN(MinSimilarity) || double.IsInfinity(MinSimilarity))
                MinSimilarity = 0.0;
            if (MinSimilarity < -1.0)
                MinSimilarity = -1.0;
            if (MinSimilarity > 1.0)
                MinSimilarity = 1.0;

            if (MaxContextChars < 1000)
                MaxContextChars = 1000;
            if (MaxContextChars > 60000)
                MaxContextChars = 60000;

            if (string.IsNullOrWhiteSpace(ChatModel))
                ChatModel = "GigaChat";
            if (string.IsNullOrWhiteSpace(EmbeddingModel))
                EmbeddingModel = "Embeddings";
            if (string.IsNullOrWhiteSpace(SourceButtonText))
                SourceButtonText = DefaultSourceButtonText;
            if (string.IsNullOrWhiteSpace(MoodleUrlTemplate))
                MoodleUrlTemplate = DefaultMoodleUrlTemplate;
        }

        [DataContract]
        private sealed class AiSearchOptionsDto
        {
            [DataMember(Name = "enableEmbeddingSearch")]
            public bool? EnableEmbeddingSearch { get; set; }

            [DataMember(Name = "topK")]
            public int? TopK { get; set; }

            [DataMember(Name = "minSimilarity")]
            public double? MinSimilarity { get; set; }

            [DataMember(Name = "maxContextChars")]
            public int? MaxContextChars { get; set; }

            [DataMember(Name = "allowAnswerWithoutContext")]
            public bool? AllowAnswerWithoutContext { get; set; }

            [DataMember(Name = "chatModel")]
            public string ChatModel { get; set; }

            [DataMember(Name = "embeddingModel")]
            public string EmbeddingModel { get; set; }

            [DataMember(Name = "sourceButtonText")]
            public string SourceButtonText { get; set; }

            [DataMember(Name = "moodleUrlTemplate")]
            public string MoodleUrlTemplate { get; set; }

            [DataMember(Name = "debugRetrieval")]
            public bool? DebugRetrieval { get; set; }
        }
    }
}