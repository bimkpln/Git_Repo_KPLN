using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace KPLN_Publication
{
    internal static class FileNameTemplate
    {
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
}
