using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using KPLN_Tools_OVVK.Forms.Entities.VentConfigurator;
using static KPLN_Tools_OVVK.Services.VentConfigurator.VentConfiguratorSettings;

namespace KPLN_Tools_OVVK.Services.VentConfigurator
{
    internal static class ProjectFamilyStorage
    {
        internal static string GetPath(ProjectItem project)
        {
            if (project == null || project.IsUnknown)
                return null;
            string code = CleanDatabaseValue(project.Code);
            if (code.Length == 0)
                throw new InvalidOperationException("В базе проектов не задана аббревиатура проекта «" + project.DisplayName + "».");
            // Аббревиатура — ровно одна папка Windows. Не подменяем её другим именем.
            if (code == "." || code == ".." || code.EndsWith(".", StringComparison.Ordinal)
                || code.Any(c => char.IsControl(c) || "<>:\"/\\|?*".IndexOf(c) >= 0)
                || Regex.IsMatch(code, @"^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)", RegexOptions.IgnoreCase))
                throw new InvalidOperationException("Аббревиатура проекта «" + code + "» содержит недопустимое имя папки.");
            return Path.Combine(ProjectFamiliesRoot, code, ProjectFamilyFileName);
        }

        internal static ProjectMatch Find(IList<ProjectItem> projects, string familyPath)
        {
            string normalized = ProjectPathMatcher.Normalize(familyPath);
            var result = new ProjectMatch();
            if (normalized.Length == 0)
                return result;
            foreach (var project in projects)
            {
                string path;
                try
                {
                    path = GetPath(project);
                }
                catch (InvalidOperationException) { continue; }
                if (path == null || !string.Equals(normalized, ProjectPathMatcher.Normalize(path), StringComparison.OrdinalIgnoreCase))
                    continue;
                if (result.Project != null && !ReferenceEquals(result.Project, project))
                    return new ProjectMatch { IsAmbiguous = true };
                result.Project = project;
                result.RootPath = Path.GetDirectoryName(path);
            }
            return result;
        }

        internal static bool EnsureFile(string sourcePath, string path, Func<string, string, bool> confirm)
        {
            if (File.Exists(path))
                return true;
            if (!confirm(sourcePath, path))
                return false;
            // Пока был открыт вопрос, файл мог появиться у другого пользователя.
            if (File.Exists(path))
                return true;
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException("Исходное семейство для копирования недоступно. Проверьте подключение к сети.", sourcePath);
            string directory = Path.GetDirectoryName(path);
            Directory.CreateDirectory(directory);
            string temporary = Path.Combine(directory, ".KPLN_Copy_" + Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                // Готовый файл появляется только после полного копирования. Чужой файл не перезаписываем.
                File.Copy(sourcePath, temporary, false);
                File.SetAttributes(temporary, File.GetAttributes(temporary) & ~FileAttributes.ReadOnly);
                try
                {
                    File.Move(temporary, path);
                }
                catch (IOException) { if (!File.Exists(path)) throw; }
                return true;
            }
            finally
            {
                try
                {
                    if (File.Exists(temporary))
                        File.Delete(temporary);
                }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            }
        }
    }
}
