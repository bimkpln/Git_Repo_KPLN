using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace KPLN_CoordiantorAI.Common
{
    public static class ArticleAliasSettings
    {
        public const string EmptyJson = "[]";

        public static IList<ArticleAliasEntry> FromJson(string json)
        {
            string normalizedJson = string.IsNullOrWhiteSpace(json) ? EmptyJson : json.Trim();
            IList<ArticleAliasEntry> entries = Deserialize<List<ArticleAliasEntry>>(normalizedJson);
            return entries ?? new List<ArticleAliasEntry>();
        }

        public static string ToJson(IEnumerable<ArticleAliasEntry> entries)
        {
            List<ArticleAliasEntry> normalizedEntries = new List<ArticleAliasEntry>();
            if (entries != null)
            {
                foreach (ArticleAliasEntry entry in entries)
                {
                    if (entry == null || string.IsNullOrWhiteSpace(entry.ArticleId) || string.IsNullOrWhiteSpace(entry.Aliases))
                        continue;

                    normalizedEntries.Add(new ArticleAliasEntry
                    {
                        ArticleId = entry.ArticleId.Trim(),
                        Aliases = NormalizeAliases(entry.Aliases)
                    });
                }
            }

            return Serialize(normalizedEntries);
        }

        public static IDictionary<string, string> ToAliasMap(IEnumerable<ArticleAliasEntry> entries)
        {
            Dictionary<string, string> aliasesByArticleId = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (entries == null)
                return aliasesByArticleId;

            foreach (ArticleAliasEntry entry in entries)
            {
                if (entry == null || string.IsNullOrWhiteSpace(entry.ArticleId))
                    continue;

                aliasesByArticleId[entry.ArticleId.Trim()] = NormalizeAliases(entry.Aliases);
            }

            return aliasesByArticleId;
        }

        public static IList<string> FindMatchingArticleIds(string queryText, IEnumerable<ArticleAliasEntry> entries)
        {
            List<string> articleIds = new List<string>();
            if (string.IsNullOrWhiteSpace(queryText) || entries == null)
                return articleIds;

            string normalizedQuery = " " + NormalizeForMatch(queryText) + " ";
            HashSet<string> knownArticleIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (ArticleAliasEntry entry in entries)
            {
                if (entry == null || string.IsNullOrWhiteSpace(entry.ArticleId) || string.IsNullOrWhiteSpace(entry.Aliases))
                    continue;

                foreach (string alias in SplitAliases(entry.Aliases))
                {
                    string normalizedAlias = NormalizeForMatch(alias);
                    if (normalizedAlias.Length < 2)
                        continue;

                    if (normalizedQuery.IndexOf(" " + normalizedAlias + " ", StringComparison.Ordinal) < 0)
                        continue;

                    string articleId = entry.ArticleId.Trim();
                    if (knownArticleIds.Add(articleId))
                        articleIds.Add(articleId);

                    break;
                }
            }

            return articleIds;
        }

        public static string NormalizeAliases(string aliases)
        {
            if (string.IsNullOrWhiteSpace(aliases))
                return string.Empty;

            List<string> normalizedAliases = new List<string>();
            HashSet<string> knownAliases = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string alias in SplitAliases(aliases))
            {
                string normalizedAlias = (alias ?? string.Empty).Trim();
                if (normalizedAlias.Length == 0 || !knownAliases.Add(normalizedAlias))
                    continue;

                normalizedAliases.Add(normalizedAlias);
            }

            return string.Join("; ", normalizedAliases.ToArray());
        }

        private static IEnumerable<string> SplitAliases(string aliases)
        {
            return (aliases ?? string.Empty).Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
        }

        private static string NormalizeForMatch(string value)
        {
            string lower = (value ?? string.Empty).Trim().ToLowerInvariant().Replace('ё', 'е');
            StringBuilder builder = new StringBuilder();
            bool previousWasSpace = true;
            foreach (char currentChar in lower)
            {
                if (char.IsLetterOrDigit(currentChar))
                {
                    builder.Append(currentChar);
                    previousWasSpace = false;
                    continue;
                }

                if (!previousWasSpace)
                {
                    builder.Append(' ');
                    previousWasSpace = true;
                }
            }

            return builder.ToString().Trim();
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
    }

    [DataContract]
    public sealed class ArticleAliasEntry
    {
        [DataMember(Name = "article_id")]
        public string ArticleId { get; set; }

        [DataMember(Name = "aliases")]
        public string Aliases { get; set; }
    }
}