using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace KPLN_CoordiantorAI.Common
{
    public static class PromptDefaults
    {
        private const string PromptsFolderName = "Prompts";

        public static string LoadSystemPrompt()
        {
            return LoadPrompt("system_prompt.txt");
        }

        public static string LoadResponseContextPrompt()
        {
            return LoadPrompt("response_context_prompt.txt");
        }

        public static string LoadArticleHintPrompt()
        {
            return LoadPrompt("article_hint_prompt.txt");
        }

        private static string LoadPrompt(string fileName)
        {
            foreach (string promptPath in GetPromptPathCandidates(fileName))
            {
                if (File.Exists(promptPath))
                    return File.ReadAllText(promptPath, Encoding.UTF8).Trim();
            }

            return string.Empty;
        }

        private static IEnumerable<string> GetPromptPathCandidates(string fileName)
        {
            string assemblyDirectory = GetAssemblyDirectory();
            if (!string.IsNullOrWhiteSpace(assemblyDirectory))
                yield return Path.Combine(assemblyDirectory, PromptsFolderName, fileName);

            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            if (!string.IsNullOrWhiteSpace(baseDirectory))
                yield return Path.Combine(baseDirectory, PromptsFolderName, fileName);

            string currentDirectory = Directory.GetCurrentDirectory();
            if (!string.IsNullOrWhiteSpace(currentDirectory))
                yield return Path.Combine(currentDirectory, PromptsFolderName, fileName);
        }

        private static string GetAssemblyDirectory()
        {
            string location = Assembly.GetExecutingAssembly().Location;
            return string.IsNullOrWhiteSpace(location) ? string.Empty : Path.GetDirectoryName(location);
        }
    }
}
