using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace KPLN_Publication
{
    internal static class FileNameTemplate
    {
        internal static string AppendSuffix(string name, string suffix)
        {
            return Path.GetFileNameWithoutExtension(name) + suffix + Path.GetExtension(name);
        }

        internal static string GetAvailableFileName(string folder, string name, ISet<string> reserved)
        {
            string candidate = name;
            int index = 1;
            while (reserved.Contains(candidate) || File.Exists(Path.Combine(folder, candidate))
                || Directory.Exists(Path.Combine(folder, candidate)))
            {
                candidate = AppendSuffix(name, "_" + index++);
            }
            reserved.Add(candidate);
            return candidate;
        }

        internal static string WithoutExportExtension(string name)
        {
            string previous;
            do
            {
                previous = name;
                name = WithoutExtension(WithoutExtension(name, ".pdf"), ".dwg");
            } while (name != previous);
            return name;
        }

        internal static string WithoutExtension(string name, string extension)
        {
            name = (name ?? string.Empty).TrimEnd();
            while (name.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
                name = name.Substring(0, name.Length - extension.Length).TrimEnd();
            return name;
        }

        internal static string WithExtension(string name, string extension)
        {
            return WithoutExtension(name, extension) + extension;
        }

        internal static string Resolve(string template, Func<string, string> getValue, out List<string> errors)
        {
            List<string> issues = new List<string>();
            template = template ?? string.Empty;
            if (Regex.Replace(template, "<[^<>]+>", string.Empty).IndexOfAny(new[] { '<', '>' }) >= 0)
                issues.Add("Непарные или пустые скобки < > в шаблоне");

            string name = Regex.Replace(template, "<([^<>]+)>", match =>
            {
                string parameterName = match.Groups[1].Value;
                string value = getValue(parameterName);
                if (value == null)
                    issues.Add("Нет параметра: " + parameterName);
                else if (string.IsNullOrWhiteSpace(value))
                    issues.Add("Пустой параметр: " + parameterName);
                return value ?? string.Empty;
            });

            // Та же очистка, что применяется при фактической выдаче.
            errors = issues.Distinct().ToList();
            return new string(name.Where(c => char.IsLetterOrDigit(c) || char.IsWhiteSpace(c)
                || c == '-' || c == '+' || c == '_' || c == '.' || c == ',').ToArray());
        }

        internal static string Validate(string name, string extension)
        {
            if (string.IsNullOrWhiteSpace(name) || name.Equals(extension, StringComparison.OrdinalIgnoreCase))
                return "Пустое имя файла";
            if (name.Any(c => c < 32 || "<>:\"/\\|?*".Contains(c)))
                return "Недопустимые символы в имени файла";
            if (name.EndsWith(".") || name.EndsWith(" "))
                return "Имя заканчивается точкой или пробелом";
            if (!name.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
                return "Нужно расширение " + extension;
            string stem = name.Split('.')[0].TrimEnd();
            if (Regex.IsMatch(stem, @"^(CON|PRN|AUX|NUL|COM[1-9¹²³]|LPT[1-9¹²³])$", RegexOptions.IgnoreCase))
                return "Зарезервированное имя Windows";
            if (name.Length > 255)
                return "Имя файла длиннее 255 символов";
            return string.Empty;
        }
    }

    // Общий план имён: готовые PDF/DWG, объединение и временные файлы обработки.
    internal sealed class ExportFilePlan
    {
        internal string Folder { get; private set; }
        internal string RequestedFolder { get; private set; }
        internal bool Redirected { get; private set; }
        internal Dictionary<string, string> Names { get; private set; }

        internal static ExportFilePlan Create(string root, string title, YayPrintSettings settings,
            string format, IEnumerable<string> requestedNames)
        {
            if (string.IsNullOrWhiteSpace(root))
                throw new ArgumentException("Не задана папка сохранения");

            List<string> names = requestedNames.ToList();
            if (names.Distinct(StringComparer.OrdinalIgnoreCase).Count() != names.Count)
                throw new InvalidOperationException("Совпадение имён выходных или промежуточных файлов");

            ExportFilePlan plan = new ExportFilePlan
            {
                RequestedFolder = Path.Combine(root, settings.GetExportFolderName(title))
            };
            try
            {
                plan.Folder = Path.GetFullPath(plan.RequestedFolder);
                plan.Names = Allocate(plan.Folder, names);
                plan.Redirected = IsTooLong(plan.Folder, plan.Names.Values);
            }
            catch (PathTooLongException)
            {
                plan.Redirected = true;
            }

            if (plan.Redirected)
            {
                string systemRoot = Path.GetPathRoot(Environment.GetFolderPath(Environment.SpecialFolder.System));
                plan.Folder = Path.Combine(systemRoot, format, settings.GetExportFallbackFolderName());
                plan.Names = Allocate(plan.Folder, names);
            }

            foreach (string name in plan.Names.Values)
                ValidatePath(plan.Folder, name);
            return plan;
        }

        private static Dictionary<string, string> Allocate(string folder, List<string> names)
        {
            // Заранее резервируем имена всего набора: A_1 не должен занять имя другого листа.
            HashSet<string> reserved = new HashSet<string>(names, StringComparer.OrdinalIgnoreCase);
            Dictionary<string, string> result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string name in names)
            {
                reserved.Remove(name);
                result.Add(name, FileNameTemplate.GetAvailableFileName(folder, name, reserved));
            }
            return result;
        }

        private static bool IsTooLong(string folder, IEnumerable<string> names)
        {
            return folder.Length >= 248 || names.Any(name => Path.Combine(folder, name).Length >= 260);
        }

        private static void ValidatePath(string folder, string name)
        {
            string error = FileNameTemplate.Validate(name, Path.GetExtension(name));
            if (!string.IsNullOrEmpty(error))
                throw new InvalidOperationException(error + ": " + name);
            if (IsTooLong(folder, new[] { name }))
                throw new PathTooLongException("Полный путь должен быть короче 260 символов: " + Path.Combine(folder, name));
        }

        internal string GetPath(string requestedName, bool forWrite = false)
        {
            string name = Names[requestedName];
            if (forWrite && (File.Exists(Path.Combine(Folder, name)) || Directory.Exists(Path.Combine(Folder, name))))
            {
                HashSet<string> reserved = new HashSet<string>(Names.Values, StringComparer.OrdinalIgnoreCase);
                name = FileNameTemplate.GetAvailableFileName(Folder, requestedName, reserved);
                ValidatePath(Folder, name);
                Names[requestedName] = name;
            }
            return Path.Combine(Folder, name);
        }
    }
}
