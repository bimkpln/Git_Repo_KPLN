using System;
using System.Collections.Generic;
using KPLN_Tools_OVVK.Forms.Entities.VentConfigurator;
using static KPLN_Tools_OVVK.Services.VentConfigurator.VentConfiguratorSettings;

namespace KPLN_Tools_OVVK.Services.VentConfigurator
{
    // Сопоставление только полных путей и границ папок: «Проект» не совпадает с «Проект2».
    internal static class ProjectPathMatcher
    {
        internal static string Normalize(string value)
        {
            string path = CleanDatabaseValue(value).Replace('\\', '/');
            string prefix;
            int protectedSegments;
            if (path.StartsWith("RSN://", StringComparison.OrdinalIgnoreCase))
            {
                prefix = "RSN://";
                path = path.Substring(6);
                protectedSegments = 1; // имя Revit Server
            }
            else if (path.StartsWith("//", StringComparison.Ordinal))
            {
                prefix = "//";
                path = path.Substring(2);
                protectedSegments = 2; // сервер и общая папка UNC
            }
            else if (path.Length >= 3 && char.IsLetter(path[0]) && path[1] == ':' && path[2] == '/')
            {
                prefix = path.Substring(0, 3);
                path = path.Substring(3);
                protectedSegments = 0;
            }
            else
                return string.Empty;

            var segments = new List<string>();
            foreach (string segment in path.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries))
            {
                if (segment == ".")
                    continue;
                if (segment == "..")
                {
                    if (segments.Count <= protectedSegments)
                        return string.Empty;
                    segments.RemoveAt(segments.Count - 1);
                }
                else
                    segments.Add(segment);
            }
            if (segments.Count < protectedSegments)
                return string.Empty;
            return prefix + string.Join("/", segments);
        }

        internal static ProjectMatch Find(IList<ProjectItem> projects, string modelPath)
        {
            string path = Normalize(modelPath);
            var match = new ProjectMatch();
            if (path.Length == 0)
                return match;
            int bestLength = -1;
            foreach (var project in projects)
            {
                foreach (string root in project.Paths ?? new string[0])
                {
                    string normalizedRoot = Normalize(root);
                    if (normalizedRoot.Length == 0)
                        continue;
                    string folderPrefix = normalizedRoot.TrimEnd('/') + "/";
                    if (!string.Equals(path, normalizedRoot, StringComparison.OrdinalIgnoreCase) &&
                        !path.StartsWith(folderPrefix, StringComparison.OrdinalIgnoreCase))
                        continue;
                    if (normalizedRoot.Length > bestLength)
                    {
                        bestLength = normalizedRoot.Length;
                        match.Project = project;
                        match.RootPath = root;
                        match.IsAmbiguous = false;
                    }
                    else if (normalizedRoot.Length == bestLength && !ReferenceEquals(match.Project, project))
                        match.IsAmbiguous = true;
                }
            }
            if (match.IsAmbiguous)
            {
                match.Project = null;
                match.RootPath = null;
            }
            return match;
        }
    }
}
