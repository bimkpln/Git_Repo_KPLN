using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace KPLN_CoordiantorAI.ExternalModel
{
    internal sealed class ModelChatLegacyLogReader
    {
        private enum EntrySection
        {
            None,
            Question,
            Answer
        }

        public IList<ModelChatHistoryEntry> ReadMissingEntries(
            string logFolder,
            IEnumerable<ModelChatHistoryEntry> databaseEntries)
        {
            if (string.IsNullOrWhiteSpace(logFolder))
                return new List<ModelChatHistoryEntry>();

            string logPath = Path.Combine(logFolder.Trim(), Environment.UserName + ".txt");
            if (!File.Exists(logPath))
                return new List<ModelChatHistoryEntry>();

            List<ModelChatHistoryEntry> databaseList =
                (databaseEntries ?? Enumerable.Empty<ModelChatHistoryEntry>()).ToList();
            HashSet<string> knownEntries = new HashSet<string>(
                databaseList.Select(BuildDuplicateKey),
                StringComparer.Ordinal);
            Dictionary<string, List<ModelIdentityCandidate>> identitiesByName =
                BuildIdentityLookup(databaseList);

            List<ModelChatHistoryEntry> result = new List<ModelChatHistoryEntry>();
            HashSet<string> acceptedEntries = new HashSet<string>(knownEntries, StringComparer.Ordinal);
            using (FileStream stream = new FileStream(
                logPath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete))
            using (StreamReader reader = new StreamReader(stream, Encoding.UTF8, true))
            {
                LegacyEntryBuilder builder = new LegacyEntryBuilder();
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (IsEntrySeparator(line))
                    {
                        TryAddEntry(builder, identitiesByName, acceptedEntries, result);
                        builder = new LegacyEntryBuilder();
                        continue;
                    }

                    builder.ReadLine(line);
                }

                TryAddEntry(builder, identitiesByName, acceptedEntries, result);
            }

            return result
                .OrderByDescending(entry => ParseDate(entry.RequestTime))
                .ToList();
        }

        private static void TryAddEntry(
            LegacyEntryBuilder builder,
            IDictionary<string, List<ModelIdentityCandidate>> identitiesByName,
            ISet<string> acceptedEntries,
            ICollection<ModelChatHistoryEntry> result)
        {
            if (builder == null
                || string.IsNullOrWhiteSpace(builder.Question)
                || string.IsNullOrWhiteSpace(builder.Answer)
                || (!string.IsNullOrWhiteSpace(builder.Scenario)
                    && !string.Equals(builder.Scenario, "wpf_window", StringComparison.OrdinalIgnoreCase)))
                return;

            DateTime requestTime = ParseDate(builder.RequestTime);
            string modelName = string.IsNullOrWhiteSpace(builder.ModelName)
                ? "Без имени"
                : builder.ModelName.Trim();
            ModelIdentityCandidate identity = ResolveIdentity(modelName, identitiesByName);
            ModelChatHistoryEntry entry = new ModelChatHistoryEntry
            {
                RequestId = string.Empty,
                SessionId = string.Empty,
                RequestTime = requestTime == DateTime.MinValue
                    ? builder.RequestTime
                    : requestTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
                UserQuestion = builder.Question.Trim(),
                FinalAnswer = builder.Answer.Trim(),
                ModelName = modelName,
                ModelIdentity = identity == null
                    ? "legacy-model-name:" + NormalizeKeyPart(modelName)
                    : identity.ModelIdentity,
                RevitVersion = identity == null ? 0 : identity.RevitVersion
            };

            string duplicateKey = BuildDuplicateKey(entry);
            if (acceptedEntries.Add(duplicateKey))
                result.Add(entry);
        }

        private static Dictionary<string, List<ModelIdentityCandidate>> BuildIdentityLookup(
            IEnumerable<ModelChatHistoryEntry> databaseEntries)
        {
            return databaseEntries
                .Where(entry => entry != null
                    && !string.IsNullOrWhiteSpace(entry.ModelName)
                    && !string.IsNullOrWhiteSpace(entry.ModelIdentity)
                    && !entry.ModelIdentity.StartsWith("legacy-session:", StringComparison.Ordinal))
                .GroupBy(entry => NormalizeKeyPart(entry.ModelName))
                .ToDictionary(
                    group => group.Key,
                    group => group
                        .Select(entry => new ModelIdentityCandidate
                        {
                            ModelIdentity = entry.ModelIdentity,
                            RevitVersion = entry.RevitVersion
                        })
                        .GroupBy(item => item.ModelIdentity + "|" + item.RevitVersion)
                        .Select(item => item.First())
                        .ToList(),
                    StringComparer.Ordinal);
        }

        private static ModelIdentityCandidate ResolveIdentity(
            string modelName,
            IDictionary<string, List<ModelIdentityCandidate>> identitiesByName)
        {
            List<ModelIdentityCandidate> candidates;
            if (!identitiesByName.TryGetValue(NormalizeKeyPart(modelName), out candidates)
                || candidates.Count != 1)
                return null;
            return candidates[0];
        }

        private static string BuildDuplicateKey(ModelChatHistoryEntry entry)
        {
            if (entry == null)
                return string.Empty;

            DateTime date = ParseDate(entry.RequestTime);
            string timePart = date == DateTime.MinValue
                ? NormalizeKeyPart(entry.RequestTime)
                : date.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            return NormalizeKeyPart(entry.ModelName)
                + "|"
                + timePart
                + "|"
                + NormalizeKeyPart(entry.UserQuestion);
        }

        private static string NormalizeKeyPart(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            StringBuilder result = new StringBuilder();
            bool previousWasWhitespace = false;
            foreach (char character in value.Trim())
            {
                if (char.IsWhiteSpace(character))
                {
                    if (!previousWasWhitespace)
                        result.Append(' ');
                    previousWasWhitespace = true;
                }
                else
                {
                    result.Append(char.ToUpperInvariant(character));
                    previousWasWhitespace = false;
                }
            }
            return result.ToString();
        }

        private static bool IsEntrySeparator(string line)
        {
            if (string.IsNullOrWhiteSpace(line))
                return false;

            string trimmed = line.Trim();
            return trimmed.Length >= 20 && trimmed.All(character => character == '-');
        }

        private static DateTime ParseDate(string value)
        {
            DateTime result;
            string[] formats =
            {
                "yyyy-MM-dd HH:mm:ss.fff",
                "yyyy-MM-dd HH:mm:ss"
            };
            if (DateTime.TryParseExact(
                value,
                formats,
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out result))
                return result;
            return DateTime.TryParse(value, CultureInfo.CurrentCulture, DateTimeStyles.None, out result)
                ? result
                : DateTime.MinValue;
        }

        private sealed class ModelIdentityCandidate
        {
            public string ModelIdentity { get; set; }
            public int RevitVersion { get; set; }
        }

        private sealed class LegacyEntryBuilder
        {
            private readonly StringBuilder _question = new StringBuilder();
            private readonly StringBuilder _answer = new StringBuilder();
            private EntrySection _section;

            public string RequestTime { get; private set; }
            public string Scenario { get; private set; }
            public string ModelName { get; private set; }
            public string Question { get { return _question.ToString(); } }
            public string Answer { get { return _answer.ToString(); } }

            public void ReadLine(string line)
            {
                string currentLine = line ?? string.Empty;
                string value;
                if (TryReadValue(currentLine, "REQUEST_TIME:", out value))
                {
                    RequestTime = value;
                    _section = EntrySection.None;
                    return;
                }
                if (TryReadValue(currentLine, "SCENARIO:", out value))
                {
                    Scenario = value;
                    _section = EntrySection.None;
                    return;
                }
                if (TryReadValue(currentLine, "REVIT_MODEL:", out value))
                {
                    ModelName = value;
                    _section = EntrySection.None;
                    return;
                }
                if (TryReadValue(currentLine, "ВОПРОС:", out value)
                    || TryReadValue(currentLine, "REQUEST:", out value))
                {
                    AppendLine(_question, value);
                    _section = EntrySection.Question;
                    return;
                }
                if (TryReadValue(currentLine, "ОТВЕТ:", out value)
                    || TryReadValue(currentLine, "RESPONSE:", out value))
                {
                    AppendLine(_answer, value);
                    _section = EntrySection.Answer;
                    return;
                }
                if (currentLine.StartsWith("--- СТАТИСТИКА", StringComparison.Ordinal)
                    || currentLine.StartsWith("--- TOOL AREA", StringComparison.Ordinal)
                    || currentLine.StartsWith("--- MCP TOOLS", StringComparison.Ordinal))
                {
                    _section = EntrySection.None;
                    return;
                }

                if (_section == EntrySection.Question)
                    AppendLine(_question, currentLine);
                else if (_section == EntrySection.Answer)
                    AppendLine(_answer, currentLine);
            }

            private static bool TryReadValue(string line, string prefix, out string value)
            {
                if (!line.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    value = null;
                    return false;
                }

                value = line.Substring(prefix.Length).TrimStart();
                return true;
            }

            private static void AppendLine(StringBuilder target, string line)
            {
                if (target.Length > 0)
                    target.AppendLine();
                target.Append(line ?? string.Empty);
            }
        }
    }
}
