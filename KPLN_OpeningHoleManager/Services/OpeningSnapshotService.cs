using Autodesk.Revit.DB;
using KPLN_OpeningHoleManager.Core;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace KPLN_OpeningHoleManager.Services
{
    internal static class OpeningSnapshotService
    {
        private static readonly string[] DimensionNames =
        {
            "ширин", "ширн", "высот", "диаметр", "радиус", "глубин", "толщин", "длин", "расширение границ",
            "width", "height", "diameter", "radius", "depth", "thickness", "length",
        };

        private static readonly string[] MainDimensionNames =
        {
            "АР_Ширна Проема", "АР_Высота Проема", "КП_Р_Высота",
            "Ширина", "Высота", "Диаметр", "Радиус", "Глубина", "Толщина",
            "Width", "Height", "Diameter", "Radius", "Depth", "Thickness",
        };

        private static string Millimeters(double value) => (value * 304.8).ToString("0.###", CultureInfo.InvariantCulture);

        internal static OpeningSnapshot Capture(FamilyInstance opening)
        {
            var snapshot = new OpeningSnapshot();
            Document document = opening.Document;
            Level level = document.GetElement(opening.LevelId) as Level;
            if (level == null && opening.Host != null)
                level = document.GetElement(opening.Host.LevelId) as Level;

            snapshot["LevelId"] = level?.UniqueId ?? string.Empty;
            snapshot["LevelName"] = level?.Name ?? "Без уровня";
            snapshot["LevelElevation"] = level == null ? string.Empty : Millimeters(level.Elevation);
            snapshot["TypeId"] = opening.Symbol.UniqueId ?? opening.GetTypeId().ToString();
            snapshot["TypeName"] = opening.Symbol.FamilyName + ": " + opening.Name;
            snapshot["HostId"] = opening.Host?.UniqueId ?? string.Empty;
            snapshot["FacingFlipped"] = opening.FacingFlipped.ToString();
            snapshot["HandFlipped"] = opening.HandFlipped.ToString();


            LocationPoint location = opening.Location as LocationPoint;
            if (location != null)
            {
                snapshot["X"] = Millimeters(location.Point.X);
                snapshot["Y"] = Millimeters(location.Point.Y);
                snapshot["Z"] = Millimeters(location.Point.Z);
                snapshot["Rotation"] = (location.Rotation * 180 / Math.PI).ToString("0.######", CultureInfo.InvariantCulture);
            }

            // Сохраняем реальные имена размерных параметров, в том числе АР_Ширна Проема из СЕТ.
            var dimensions = new SortedDictionary<string, string>(StringComparer.Ordinal);
            ReadDimensions(opening, "Экземпляр", dimensions);
            ReadDimensions(opening.Symbol, "Тип", dimensions);
            snapshot.Dimensions = dimensions;

            return snapshot;
        }

        private static void ReadDimensions(Element element, string scope, IDictionary<string, string> target)
        {
            foreach (Parameter parameter in element.Parameters)
            {
                string name = parameter.Definition?.Name;
                if (parameter.StorageType != StorageType.Double || string.IsNullOrEmpty(name)
                    || !DimensionNames.Any(token => name.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0))
                    continue;

                target[scope + ": " + name + " [" + parameter.Id + "]"] = Millimeters(parameter.AsDouble());
            }
        }


        internal static string Describe(OpeningSnapshot snapshot)
        {
            if (snapshot == null)
                return DescribeGeometry(null);

            return "Тип: " + Value(snapshot["TypeName"]) + Environment.NewLine
                + "Уровень: " + LevelText(snapshot) + Environment.NewLine
                + DescribeGeometry(snapshot);
        }

        internal static string DescribeGeometry(OpeningSnapshot snapshot)
        {
            if (snapshot == null)
                return "Данные ещё не записаны. Нажмите «Обновить».";

            var lines = new List<string>
            {
                "Точка размещения, мм:",
                "X: " + Value(snapshot["X"]),
                "Y: " + Value(snapshot["Y"]),
                "Z: " + Value(snapshot["Z"]),
                "Поворот, °: " + Value(snapshot["Rotation"]),
            };


            KeyValuePair<string, string>[] dimensions = snapshot.Dimensions.ToArray();
            var mainDimensions = dimensions.Where(dimension => dimension.Key.StartsWith("Экземпляр: ", StringComparison.Ordinal)
                && MainDimensionNames.Contains(ParameterName(dimension.Key), StringComparer.OrdinalIgnoreCase)).ToArray();

            // У СЕТ рядом с фактическими размерами экземпляра есть нулевые размеры типоразмера.
            // В карточке показываем основные размеры отверстия; полный набор остаётся в снимках для сравнения.
            if (mainDimensions.Length > 0)
                dimensions = mainDimensions;

            lines.Add(dimensions.Length == 0 ? "Размеры: не записаны" : "Размеры, мм:");
            foreach (KeyValuePair<string, string> dimension in dimensions)
            {
                // ID параметра нужен для сравнения, но не для чтения истории пользователем.
                string name = ParameterName(dimension.Key);
                if (dimension.Key.StartsWith("Тип: ", StringComparison.Ordinal))
                    name += " (тип)";

                lines.Add(name + ": " + Value(dimension.Value));
            }

            return string.Join(Environment.NewLine, lines);
        }

        private static string ParameterName(string key)
        {
            int scopeEnd = key.IndexOf(": ", StringComparison.Ordinal);
            string name = scopeEnd < 0 ? key : key.Substring(scopeEnd + 2);
            int idStart = name.LastIndexOf(" [", StringComparison.Ordinal);

            return idStart >= 0 && name.EndsWith("]", StringComparison.Ordinal) ? name.Substring(0, idStart) : name;
        }


        internal static string[] Compare(OpeningSnapshot previous, OpeningSnapshot current)
        {
            var changes = new List<string>();
            if (new[] { "LevelId", "LevelName", "LevelElevation" }.Any(field => previous[field] != current[field]))
                changes.Add("Уровень: " + LevelText(previous) + " → " + LevelText(current));

            if (new[] { "X", "Y", "Z" }.Any(field => previous[field] != current[field]))
                changes.Add("Точка размещения, мм: " + PointText(previous) + " → " + PointText(current));

            if (previous["Rotation"] != current["Rotation"])
                changes.Add("Поворот, °: " + Value(previous["Rotation"]) + " → " + Value(current["Rotation"]));

            if (previous["TypeId"] != current["TypeId"] || previous["TypeName"] != current["TypeName"])
                changes.Add("Тип: " + previous["TypeName"] + " → " + current["TypeName"]);

            if (previous["HostId"] != current["HostId"])
                changes.Add("Основа: " + Value(previous["HostId"]) + " → " + Value(current["HostId"]));

            if (previous["FacingFlipped"] != current["FacingFlipped"] || previous["HandFlipped"] != current["HandFlipped"])
                changes.Add("Изменено отражение отверстия");


            SortedDictionary<string, string> oldSizes = previous.Dimensions;
            SortedDictionary<string, string> newSizes = current.Dimensions;
            foreach (string key in oldSizes.Keys.Union(newSizes.Keys).OrderBy(value => value, StringComparer.Ordinal))
            {
                oldSizes.TryGetValue(key, out string before);
                newSizes.TryGetValue(key, out string after);
                if (before != after)
                    changes.Add(key + ", мм: " + Value(before) + " → " + Value(after));
            }

            return changes.ToArray();
        }

        private static string Value(string text) => string.IsNullOrEmpty(text) ? "не задано" : text;
        private static string LevelText(OpeningSnapshot item) => item["LevelName"] + " (" + Value(item["LevelElevation"]) + " мм)";
        private static string PointText(OpeningSnapshot item) => "X=" + Value(item["X"]) + "; Y=" + Value(item["Y"]) + "; Z=" + Value(item["Z"]);
    }
}
