using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace KPLN_Tools_OVVK.Common.OVVK_System
{
    internal static class SystemNameValidation
    {
        internal static string[][] FindDuplicateNames(System.Collections.Generic.IEnumerable<string> names, string separator)
        {
            // Сравниваем состав группы без учёта порядка и пробелов вокруг имён.
            // Длина каждой части сохраняет однозначность ключа даже для нестандартных разделителей.
            string delimiter = string.IsNullOrEmpty(separator) ? "/" : separator;
            return names.Distinct(StringComparer.Ordinal)
                .GroupBy(name => string.Concat(name.Split(new[] { delimiter }, StringSplitOptions.None)
                    .Select(part => part.Trim()).OrderBy(part => part, StringComparer.Ordinal)
                    .Select(part => part.Length + ":" + part)), StringComparer.Ordinal)
                .Select(group => group.OrderBy(name => name, StringComparer.Ordinal).ToArray())
                .Where(group => group.Length > 1).ToArray();
        }

        internal static bool IsServiceValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;

            string text = Regex.Replace(value.ToLowerInvariant().Replace('ё', 'е'), @"\s+", " ").Trim();

            // Проверяем смысл служебных подсказок, сохраняя обычные имена вроде «Ручная вентиляция».
            bool manual = Regex.IsMatch(text, @"\bвручную\b|\bручн\w*\s+(ввод\w*|заполн\w*)");
            bool dynamo = Regex.IsMatch(text, @"\b(dynamo|динамо)\b");
            bool entryInstruction = Regex.IsMatch(text, @"\b(задается|задать|задайте|задавать|заполн\w*|ввод\w*|ввести|указать|укажите|указыва\w*)\b");
            if ((manual || dynamo) && entryInstruction)
                return true;

            bool transferInstruction = Regex.IsMatch(text, @"\b(значени\w*|имя\s+систем\w*|параметр\w*)\b")
                && Regex.IsMatch(text, @"\b(перенос\w*|скопир\w*|копиру\w*|не\s+(сбива\w*|перезапис\w*))\b");
            if (transferInstruction)
                return true;

            // Проверяем и составные имена: «П1 / Нет системы» тоже содержит ошибку.
            return text.Split(new[] { '/', ';', ',', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(part => part.Trim(' ', '.', ':', '!', '<', '>', '«', '»', '"'))
                .Any(part => Regex.IsMatch(part,
                    @"^(нет\s+систем\w*|без\s+систем\w*|не\s+(задан[ао]?|определен[ао]?)|систем\w*\s+(нет|отсутствует|не\s+(задана|определена|назначена))|(имя\s+)?систем\w*\s+не\s+указан[ао]?)$"));
        }
    }
}
