using System;
using System.Linq;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace KPLN_ViewsAndLists_Ribbon.Common.Lists
{
    // Формулы существующей команды; расчёт выполняется без изменения документа Revit.
    internal static class ListRenumberNumberCalculator
    {
        internal static string UserNumber(string number)
        {
            char[] charArray = number.ToCharArray();
            foreach (char c in charArray)
            {
                if (Char.IsLetter(c))
                    number = number.Trim(c);
                else if (Enum.IsDefined(typeof(UniDecCodes), (int)c))
                    number = number.Trim(c);
                else if (c.Equals('.'))
                    continue;
                else if (Char.IsPunctuation(c))
                    number = number.Split(c)[1];
            }

            if (number.All(c => c.Equals('0')))
                return "0";
            else
                number = number.Length > 1 ? number.TrimStart('0') : number;

            return number;
        }

        private static int OnlyNumber(string number)
        {
            char[] charArray = number.ToCharArray();
            if (number.Contains('/'))
            {
                string[] splitedNumber = number.Split('/');
                string lastPartSplNumber = splitedNumber[splitedNumber.Length - 1];
                charArray = lastPartSplNumber.ToCharArray();

                foreach (char c in charArray)
                {
                    if (c.Equals('.'))
                    {
                        number = number.Split('.')[0];
                        break;
                    }
                }

                return Int32.Parse(new String(lastPartSplNumber.Where(Char.IsDigit).ToArray()));
            }

            foreach (char c in charArray)
            {
                if (c.Equals('.'))
                {
                    number = number.Split('.')[0];
                    break;
                }
            }

            return Int32.Parse(new String(number.Where(Char.IsDigit).ToArray()));
        }

        internal static string Prefix(string sheetNumber, string prefix)
        {
            string onlyNumb = UserNumber(sheetNumber);
            string zeroNumber;
            if (OnlyNumber(onlyNumb) < 10)
                zeroNumber = $"00{onlyNumb}";
            else if (OnlyNumber(onlyNumb) < 100)
                zeroNumber = $"0{onlyNumb}";
            else
                zeroNumber = onlyNumb;

            return prefix + "/" + zeroNumber;
        }

        internal static string Renumber(string sheetNumber, int number)
        {
            string textPrefix = string.Empty;
            string zeroNumberPrefix = string.Empty;
            string subNumberOfNumber = string.Empty;

            var matchNumberByPrefix = Regex.Match(sheetNumber, @"([А-ЯA-Z]+)/");
            var matchNumberByPrefixWithIntNumb = Regex.Match(sheetNumber, @"([А-ЯA-Z]+)\d/");
            var matchNumberByPrefixWithDoubleNumb = Regex.Match(sheetNumber, @"([А-ЯA-Z]+)(\d?)\.(\d?)/");
            var matchByZeroPrefix = Regex.Match(sheetNumber, @"(/[0]?)");
            var matchBySubNumber = Regex.Match(sheetNumber, @"/(\d?)\.(\d?)");

            if (matchNumberByPrefix.Success)
                textPrefix = matchNumberByPrefix.Value;
            else if (matchNumberByPrefixWithIntNumb.Success)
                textPrefix = matchNumberByPrefixWithIntNumb.Value;
            else if (matchNumberByPrefixWithDoubleNumb.Success)
                textPrefix = matchNumberByPrefixWithDoubleNumb.Value;

            if (matchByZeroPrefix.Success && matchByZeroPrefix.Value.Contains("0"))
            {
                if (number < 10)
                    zeroNumberPrefix = "00";
                else if (number < 100)
                    zeroNumberPrefix = "0";
            }

            if (matchBySubNumber.Success)
                subNumberOfNumber = $".{matchBySubNumber.Value.Split('.')[1]}";
            return $"{textPrefix}{zeroNumberPrefix}{number}{subNumberOfNumber}";
        }
    }

    // Отдельный разбор для ручного продолжения: исходные формулы режимов не меняются.
    internal sealed class ListRenumberSequenceNumber
    {
        private static readonly Regex NumberPattern = new Regex(@"^(?<prefix>[\p{L}_][\p{L}_ \-]*)?(?<number>[0-9]+)(?<suffix>(?:\.[0-9]+)*(?:[\p{L}]+)?)$");
        private readonly string _original;
        private readonly int _start;
        private readonly int _length;
        private readonly int _padding;

        private ListRenumberSequenceNumber(string original, int start, int length, long value, int padding)
        {
            _original = original;
            _start = start;
            _length = length;
            Value = value;
            _padding = padding;
        }

        internal long Value { get; }

        internal static string NumberPart(string fullNumber)
        {
            string visible = VisibleNumber(fullNumber, out List<int> positions);
            return visible.Substring(NumberOffset(visible));
        }

        internal static string ReplaceNumberPart(string fullNumber, string value)
        {
            fullNumber = fullNumber ?? string.Empty;
            string visible = VisibleNumber(fullNumber, out List<int> positions);
            int offset = NumberOffset(visible);
            if (offset == visible.Length) return fullNumber + value;

            var result = new StringBuilder(fullNumber.Substring(0, positions[offset]));
            int originalLength = visible.Length - offset;
            // Скрытые символы между цифрами тоже остаются в строке и сохраняют порядок.
            for (int index = 0; index < Math.Max(originalLength, value.Length); index++)
            {
                if (index > 0 && index < originalLength)
                {
                    int gapStart = positions[offset + index - 1] + 1;
                    result.Append(fullNumber, gapStart, positions[offset + index] - gapStart);
                }
                if (index < value.Length) result.Append(value[index]);
            }
            result.Append(fullNumber.Substring(positions[positions.Count - 1] + 1));
            return result.ToString();
        }

        private static int NumberOffset(string visible)
        {
            int slash = visible.LastIndexOf('/');
            if (slash >= 0) return slash + 1;
            Match match = NumberPattern.Match(visible);
            return match.Success ? match.Groups["prefix"].Length : 0;
        }

        internal static string Prefix(string number)
        {
            string visible = VisibleNumber(number, out List<int> positions);
            int slash = visible.LastIndexOf('/');
            if (slash >= 0)
            {
                string prefixBeforeSeparator = visible.Substring(0, slash);
                return string.IsNullOrEmpty(prefixBeforeSeparator) ? "—" : prefixBeforeSeparator;
            }

            Match match = NumberPattern.Match(visible);
            string prefix = match.Success ? match.Groups["prefix"].Value : string.Empty;
            return string.IsNullOrEmpty(prefix) ? "—" : prefix;
        }

        internal static ListRenumberSequenceNumber Parse(string number)
        {
            string visible = VisibleNumber(number, out List<int> positions);
            int offset = visible.LastIndexOf('/') + 1;
            Match match = NumberPattern.Match(visible.Substring(offset));
            Group digits = match.Groups["number"];
            if (!match.Success || (offset > 0 && match.Groups["prefix"].Length > 0)
                || !long.TryParse(digits.Value, NumberStyles.None, CultureInfo.InvariantCulture, out long value))
                throw new FormatException("Нужен целый номер, например 7, 007, АР/007 или АР/007.1.");

            int first = positions[offset + digits.Index];
            int last = positions[offset + digits.Index + digits.Length - 1];
            // Не перемещаем скрытые символы, если они находятся между цифрами.
            if (last - first + 1 != digits.Length)
                throw new FormatException("Между цифрами номера есть скрытые символы. Исправьте эту строку вручную.");

            return new ListRenumberSequenceNumber(number, first, digits.Length, value,
                digits.Value[0] == '0' ? digits.Length : 0);
        }

        internal string WithValue(long value)
        {
            string digits = value.ToString(CultureInfo.InvariantCulture);
            if (_padding > 0) digits = digits.PadLeft(_padding, '0');
            return _original.Substring(0, _start) + digits + _original.Substring(_start + _length);
        }

        private static string VisibleNumber(string number, out List<int> positions)
        {
            number = number ?? string.Empty;
            positions = new List<int>();
            var visible = new StringBuilder();
            for (int index = 0; index < number.Length; index++)
            {
                bool pair = char.IsHighSurrogate(number[index]) && index + 1 < number.Length
                    && char.IsLowSurrogate(number[index + 1]);
                int length = pair ? 2 : 1;
                int code = pair ? char.ConvertToUtf32(number[index], number[index + 1]) : number[index];
                UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(number, index);
                bool hidden = Enum.IsDefined(typeof(UniDecCodes), code)
                    || category == UnicodeCategory.Format || category == UnicodeCategory.Control
                    || category == UnicodeCategory.LineSeparator || category == UnicodeCategory.ParagraphSeparator
                    || category == UnicodeCategory.SpaceSeparator && code != 32;
                if (!hidden)
                {
                    visible.Append(number, index, length);
                    for (int part = 0; part < length; part++) positions.Add(index + part);
                }
                index += length - 1;
            }
            return visible.ToString();
        }
    }
}
