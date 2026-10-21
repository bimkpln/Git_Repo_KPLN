using Autodesk.Revit.DB;
using KPLN_ViewsAndLists_Ribbon.Common.Lists;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Runtime.CompilerServices;

namespace KPLN_ViewsAndLists_Ribbon.Forms.Entities
{
    public sealed class ListRenumberPreviewRow : INotifyPropertyChanged
    {
        private string _newNumber;
        private string _editableNumber;
        private string _newParameterValue;
        private string _error;

        internal ListRenumberPreviewRow(ViewSheet sheet, Parameter parameter)
        {
            Sheet = sheet;
            TargetParameter = parameter;
            CanWriteParameter = parameter != null && !parameter.IsReadOnly && parameter.StorageType == StorageType.String;
            SheetName = sheet.Name;
            OriginalNumber = sheet.SheetNumber;
            _newNumber = OriginalNumber;
            _editableNumber = ListRenumberSequenceNumber.NumberPart(OriginalNumber);
            OriginalParameterValue = parameter != null && parameter.StorageType == StorageType.String
                ? parameter.AsString() ?? string.Empty : string.Empty;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        internal ViewSheet Sheet { get; }
        internal Parameter TargetParameter { get; set; }
        internal bool CanWriteParameter { get; set; }
        internal string ModelError { get; set; }
        internal bool UpdateParameter { get; set; }
        internal string CalculationError { get; set; }
        internal string NumberEditError { get; private set; }
        internal string ParameterCalculationError { get; private set; }

        public string SheetName { get; }
        public string OriginalNumber { get; }
        public string OriginalNumberText => ListRenumberSequenceNumber.NumberPart(OriginalNumber);
        public string OriginalParameterValue { get; }
        public string OriginalPrefix => ListRenumberSequenceNumber.Prefix(OriginalNumber);
        public string NewPrefix => ListRenumberSequenceNumber.Prefix(NewNumber);
        public string OriginalUnicode => DescribeUnicode(OriginalNumber, false);
        public string NewUnicode => DescribeUnicode(NewNumber, false);
        public string OriginalUnicodeDetails => DescribeUnicode(OriginalNumber, true);
        public string NewUnicodeDetails => DescribeUnicode(NewNumber, true);

        public string NewNumber
        {
            get => _newNumber;
            internal set
            {
                if (_newNumber == value) return;

                _newNumber = value;
                _editableNumber = ListRenumberSequenceNumber.NumberPart(value);
                NumberEditError = null;
                CalculationError = null;
                RefreshParameterValue();
                OnPropertyChanged();
                OnPropertyChanged(nameof(NewUnicode));
                OnPropertyChanged(nameof(NewUnicodeDetails));
                OnPropertyChanged(nameof(NewPrefix));
                OnPropertyChanged(nameof(EditableNumber));
            }
        }

        public string EditableNumber
        {
            get => _editableNumber;
            set
            {
                _editableNumber = value ?? string.Empty;
                if (!Regex.IsMatch(_editableNumber, @"\A[0-9]+(?:\.[0-9]+)*\z"))
                {
                    NumberEditError = "Введите только номер, например 003 или 003.1. Приставки и Unicode задаются в предыдущем окне.";
                }
                else
                {
                    NumberEditError = null;
                    NewNumber = ListRenumberSequenceNumber.ReplaceNumberPart(NewNumber, _editableNumber);
                    CalculationError = null;
                }
                OnPropertyChanged();
            }
        }

        public string NewParameterValue
        {
            get => _newParameterValue;
            set
            {
                _newParameterValue = value;
                ParameterCalculationError = null;
                OnPropertyChanged();
            }
        }

        public string Error
        {
            get => _error;
            internal set
            {
                if (_error == value) return;
                _error = value;
                OnPropertyChanged();
            }
        }

        internal void RefreshParameterValue()
        {
            if (!UpdateParameter) return;

            try
            {
                string number = ListRenumberNumberCalculator.UserNumber(NewNumber ?? string.Empty);
                _newParameterValue = int.TryParse(NewNumber, out int integer) ? integer.ToString() : number;
                ParameterCalculationError = null;
            }
            catch (Exception)
            {
                _newParameterValue = OriginalParameterValue;
                ParameterCalculationError = "Не удалось рассчитать параметр. Введите его значение вручную.";
            }

            OnPropertyChanged(nameof(NewParameterValue));
        }

        // Читаем реальные символы номера, а не служебный параметр листа.
        private static string DescribeUnicode(string number, bool includeCodes)
        {
            var order = new List<int>();
            var counts = new Dictionary<int, int>();
            for (int index = 0; index < (number ?? string.Empty).Length; index++)
            {
                bool pair = char.IsHighSurrogate(number[index]) && index + 1 < number.Length
                    && char.IsLowSurrogate(number[index + 1]);
                int code = pair ? char.ConvertToUtf32(number[index], number[index + 1]) : number[index];
                UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(number, index);
                if (pair) index++;

                if (!Enum.IsDefined(typeof(UniDecCodes), code)
                    && category != UnicodeCategory.Format && category != UnicodeCategory.Control
                    && category != UnicodeCategory.LineSeparator && category != UnicodeCategory.ParagraphSeparator
                    && !(category == UnicodeCategory.SpaceSeparator && code != 32))
                    continue;

                if (!counts.ContainsKey(code))
                {
                    order.Add(code);
                    counts.Add(code, 0);
                }
                counts[code]++;
            }

            var descriptions = new List<string>();
            foreach (int code in order)
            {
                string name;
                switch (code)
                {
                    case 0x00A0: name = "NBSP"; break;
                    case 0x200B: name = "ZWSP"; break;
                    case 0x200C: name = "ZWNJ"; break;
                    case 0x202A: name = "LRE"; break;
                    case 0x202B: name = "RLE"; break;
                    case 0x202C: name = "PDF"; break;
                    case 0x202D: name = "LRO"; break;
                    case 0x202E: name = "RLO"; break;
                    case 0x2060: name = "WJ"; break;
                    case 0x2066: name = "LRI"; break;
                    case 0x2067: name = "RLI"; break;
                    case 0x2068: name = "FSI"; break;
                    case 0x2069: name = "PDI"; break;
                    case 0xFEFF: name = "BOM"; break;
                    default: name = Enum.GetName(typeof(UniDecCodes), code); break;
                }
                string codeText = "U+" + code.ToString("X4", CultureInfo.InvariantCulture);
                string description = name == null ? codeText : includeCodes ? $"{name} ({codeText})" : name;
                if (counts[code] > 1) description += $" × {counts[code]}";
                descriptions.Add(description);
            }

            return descriptions.Count == 0 ? "—" : string.Join("; ", descriptions);
        }

        private void OnPropertyChanged([CallerMemberName] string name = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
