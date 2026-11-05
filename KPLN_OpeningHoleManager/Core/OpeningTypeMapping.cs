using Autodesk.Revit.DB;
using KPLN_OpeningHoleManager.Core.MainEntity;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_OpeningHoleManager.Core
{
    // В Shared-конфигурации сохраняются имена, а не ElementId из другой модели проекта.
    public sealed class OpeningTypeAssignment
    {
        public string Department { get; set; }
        public string RectangleFamily { get; set; }
        public string RectangleType { get; set; }
        public string CircleFamily { get; set; }
        public string CircleType { get; set; }
    }

    public sealed class OpeningTypeMappingRow : INotifyPropertyChanged
    {
        private string _rectangleType;
        private string _circleType;

        public string Department { get; }
        public string[] RectangleTypes { get; }
        public string[] CircleTypes { get; }
        public event PropertyChangedEventHandler PropertyChanged;

        public OpeningTypeMappingRow(string department, string[] rectangleTypes, string[] circleTypes)
        {
            Department = department;
            RectangleTypes = rectangleTypes;
            CircleTypes = circleTypes;
        }

        public string RectangleType
        {
            get => _rectangleType;
            set { _rectangleType = RectangleTypes.Contains(value) ? value : null; Notify(); }
        }

        public string CircleType
        {
            get => _circleType;
            set { _circleType = CircleTypes.Contains(value) ? value : null; Notify(); }
        }

        private void Notify([CallerMemberName] string name = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public sealed class OpeningTypeMapping : INotifyPropertyChanged
    {
        public static readonly string[] Departments = { "Общее", "ОВ", "ВК", "ПТ", "ЭОМ", "СС" };
        private Document _document;
        private string _rectangleFamily;
        private string _circleFamily;

        public ObservableCollection<OpeningTypeMappingRow> Rows { get; } = new ObservableCollection<OpeningTypeMappingRow>();
        public event PropertyChangedEventHandler PropertyChanged;
        public string FamilyDescription => "Прямоугольные: " + _rectangleFamily + "\nКруглые: " + _circleFamily;
        private bool _isExpanded = true;
        public bool HasMissingAssignments => Rows.Count == 0 || Rows.Any(r =>
            string.IsNullOrEmpty(r.RectangleType) || string.IsNullOrEmpty(r.CircleType));

        public bool IsExpanded
        {
            get => _isExpanded;
            set
            {
                _isExpanded = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsExpanded)));
            }
        }

        public string Status
        {
            get
            {
                if (Rows.Count == 0)
                    return "Откройте настройки и назначьте типы отверстий для всех разделов.";
                if (Rows.Any(r => r.RectangleTypes.Length == 0 || r.CircleTypes.Length == 0))
                    return "В модели отсутствуют нужные семейства. Нажмите «Загрузить семейства».";
                int missing = Rows.Sum(r => (string.IsNullOrEmpty(r.RectangleType) ? 1 : 0) + (string.IsNullOrEmpty(r.CircleType) ? 1 : 0));
                return missing == 0 ? "Все типы назначены. Команды доступны." : "Назначьте типы для обеих форм отверстий. Не заполнено: " + missing + ".";
            }
        }

        public static string GetFamilyPath(Document doc, bool round)
        {
            if (doc.Title.Contains("СЕТ_1"))
                return round
                    ? @"X:\BIM\3_Семейства\8_Библиотека семейств Самолета\01_АР\Отверстия, шахты, решетки\ASML_АР_Отверстие круглое.rfa"
                    : @"X:\BIM\3_Семейства\8_Библиотека семейств Самолета\01_АР\Отверстия, шахты, решетки\ASML_АР_Отверстие прямоугольное.rfa";

            return round
                ? @"X:\BIM\3_Семейства\1_АР\199_Отверстия, ниши и шахты\199_Отверстие круглое_(Об_Стена).rfa"
                : @"X:\BIM\3_Семейства\1_АР\199_Отверстия, ниши и шахты\199_Отверстие прямоугольное_(Об_Стена).rfa";
        }

        public static string GetFamilyName(Document doc, bool round) => Path.GetFileNameWithoutExtension(GetFamilyPath(doc, round));

        private static FamilySymbol[] GetSymbols(Document doc, string family) => new FilteredElementCollector(doc)
            .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
            .Where(s => s.FamilyName.Equals(family, StringComparison.Ordinal))
            .OrderBy(s => s.Name, StringComparer.OrdinalIgnoreCase).ThenBy(s => s.Name, StringComparer.Ordinal).ToArray();

        public void Load(Document doc, IEnumerable<OpeningTypeAssignment> saved)
        {
            _document = doc;
            _rectangleFamily = GetFamilyName(doc, false);
            _circleFamily = GetFamilyName(doc, true);
            string[] rectangle = GetSymbols(doc, _rectangleFamily).Select(s => s.Name).ToArray();
            string[] circle = GetSymbols(doc, _circleFamily).Select(s => s.Name).ToArray();
            OpeningTypeAssignment[] values = saved?.ToArray() ?? new OpeningTypeAssignment[0];

            foreach (OpeningTypeMappingRow old in Rows)
                old.PropertyChanged -= RowChanged;
            Rows.Clear();
            foreach (string department in Departments)
            {
                OpeningTypeMappingRow row = new OpeningTypeMappingRow(department, rectangle, circle);
                OpeningTypeAssignment value = values.FirstOrDefault(v => v != null && v.Department == department);
                if (value != null)
                {
                    if (value.RectangleFamily == _rectangleFamily && rectangle.Contains(value.RectangleType))
                        row.RectangleType = value.RectangleType;
                    if (value.CircleFamily == _circleFamily && circle.Contains(value.CircleType))
                        row.CircleType = value.CircleType;
                }
                row.PropertyChanged += RowChanged;
                Rows.Add(row);
            }
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(FamilyDescription)));
            RowChanged(this, null);
            IsExpanded = HasMissingAssignments;
        }

        private void RowChanged(object sender, PropertyChangedEventArgs e)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Status)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasMissingAssignments)));
            if (HasMissingAssignments)
                IsExpanded = true;
        }

        public OpeningTypeAssignment[] Export() => Rows.Select(r => new OpeningTypeAssignment
        {
            Department = r.Department,
            RectangleFamily = _rectangleFamily,
            RectangleType = r.RectangleType,
            CircleFamily = _circleFamily,
            CircleType = r.CircleType
        }).ToArray();

        public bool IsForDocument(Document doc) => _document != null && _document.IsValidObject && _document.Equals(doc);

        public void Refresh(Document doc) => Load(doc, IsForDocument(doc) ? Export() : null);

        public static bool MatchesTypeName(string name, string department)
        {
            if (string.IsNullOrWhiteSpace(name))
                return false;
            string[] tokens = department == "Общее" ? new[] { "Общее", "Общий", "Несколько категорий" }
                : department == "ОВ" ? new[] { "ОВ", "ИТП" }
                : department == "СС" ? new[] { "СС", "АВ" }
                : new[] { department };
            return tokens.Any(token => name.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        public void AutoDetect() => AutoDetect(false);

        public void AutoDetect(bool onlyEmpty)
        {
            foreach (OpeningTypeMappingRow row in Rows)
            {
                if (!onlyEmpty || string.IsNullOrEmpty(row.RectangleType))
                    row.RectangleType = row.RectangleTypes.FirstOrDefault(t => MatchesTypeName(t, row.Department));
                if (!onlyEmpty || string.IsNullOrEmpty(row.CircleType))
                    row.CircleType = row.CircleTypes.FirstOrDefault(t => MatchesTypeName(t, row.Department));
            }
        }

        public static string NormalizeDepartment(string code)
        {
            switch ((code ?? "").Trim().ToUpperInvariant())
            {
                case "ОВ": case "ОВИК": case "ИТП": return "ОВ";
                case "ВК": return "ВК";
                case "ПТ": return "ПТ";
                case "ЭОМ": return "ЭОМ";
                case "СС": case "АВ": return "СС";
                default: return "Общее";
            }
        }

        public string DepartmentForSymbol(FamilySymbol symbol)
        {
            string[] matches = Rows.Where(r => (symbol.FamilyName == _rectangleFamily && r.RectangleType == symbol.Name)
                || (symbol.FamilyName == _circleFamily && r.CircleType == symbol.Name)).Select(r => r.Department).ToArray();
            if (matches.Length == 1)
                return matches[0];
            // При одном типе для нескольких разделов их нельзя достоверно различить.
            if (matches.Length > 1)
                return "Общее";
            return Departments.Skip(1).FirstOrDefault(d => MatchesTypeName(symbol.Name, d)) ?? "Общее";
        }

        public bool Validate(Document doc, out string message)
        {
            if (!IsForDocument(doc))
            {
                message = "Откройте настройки сопоставления для текущей модели.";
                return false;
            }

            string[] rectangle = GetSymbols(doc, _rectangleFamily).Select(s => s.Name).ToArray();
            string[] circle = GetSymbols(doc, _circleFamily).Select(s => s.Name).ToArray();
            string[] missing = Departments.Where(d => !Rows.Any(r => r.Department == d
                && rectangle.Contains(r.RectangleType) && circle.Contains(r.CircleType))).ToArray();
            message = missing.Length == 0 ? string.Empty
                : "Настройте сопоставление типов отверстий для разделов: " + string.Join(", ", missing)
                    + ".\nОткройте «Настройки расстановки» → «Сопоставление типов». Назначьте прямоугольные и круглые типы.";
            return missing.Length == 0;
        }

        public OpeningTypeMapping Snapshot(Document doc)
        {
            OpeningTypeMapping result = new OpeningTypeMapping();
            result.Load(doc, Export());
            return result;
        }

        internal FamilySymbol Resolve(Document doc, OpenigHoleShape shape, string department)
        {
            if (!IsForDocument(doc))
                throw new InvalidOperationException("Модель изменилась. Повторите команду из нужной модели.");
            OpeningTypeMappingRow row = Rows.FirstOrDefault(r => r.Department == NormalizeDepartment(department));
            bool round = shape == OpenigHoleShape.Round;
            string name = round ? row?.CircleType : row?.RectangleType;
            FamilySymbol symbol = GetSymbols(doc, round ? _circleFamily : _rectangleFamily).FirstOrDefault(s => s.Name == name)
                ?? throw new InvalidOperationException("Назначенный тип отверстия отсутствует. Настройте сопоставление типов.");

            if (!symbol.IsActive)
                symbol.Activate();
            return symbol;
        }
    }
}
