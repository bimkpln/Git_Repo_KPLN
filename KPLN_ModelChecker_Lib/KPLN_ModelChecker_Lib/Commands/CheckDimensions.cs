using Autodesk.Revit.DB;
using KPLN_ModelChecker_Lib.Common;
using KPLN_ModelChecker_Lib.Core;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace KPLN_ModelChecker_Lib.Commands
{
    internal class DimTypeCompare : IEqualityComparer<DimensionType>
    {
        public bool Equals(DimensionType x, DimensionType y) => x.Id.Equals(y.Id);

        public int GetHashCode(DimensionType obj) => obj.Id.GetHashCode();
    }

    public sealed class CheckDimensions : AbstrCheck
    {
        /// <summary>
        /// Список сепараторов, для поиска диапозона у размеров ПО МАКСИМУМУ
        /// </summary>
        private static readonly string[] _separArr_MAX = new string[]
        {
            "max",
            "макс",
        };

        /// <summary>
        /// Список сепараторов, для поиска диапозона у размеров ПО МИНИМУМУ
        /// </summary>
        private static readonly string[] _separArr_MIN = new string[]
        {
            "min",
            "мин",
        };

        /// <summary>
        /// Список сепараторов, для поиска диапозона у размеров
        /// </summary>
        private static readonly string[] _separArr = new string[]
        {
            "...",
            "до",
            "-",
            _separArr_MAX[0],
            _separArr_MAX[1],
            _separArr_MIN[0],
            _separArr_MIN[1],
        };

        public CheckDimensions() : base()
        {
            if (PluginName == null)
                PluginName = "Проверка размеров";

            if (ESEntity == null)
                ESEntity = new ExtensibleStorageEntity(
                    PluginName,
                    "KPLN_CheckDimensions",
                    new Guid("f2e615e0-a15b-43df-a199-a88d18a2f568"),
                    new Guid("f2e615e0-a15b-43df-a199-a88d18a2f569"));
        }

        public override Element[] GetElemsToCheck()
        {
            List<Element> result = new List<Element>();

            Element[] docDimensions = new FilteredElementCollector(CheckDocument)
                .OfClass(typeof(Dimension))
                .WhereElementIsNotElementType()
                .ToArray();

            // Анализирую ТОЛЬКО размеры на листах
            foreach (Element elem in docDimensions)
            {
                Dimension dim = (Dimension)elem;
                View dimView = dim.View;
                // Такое может быть, т.к. этого же класса зависимости (сам размер может быть удален)
                if (dimView == null)
                    continue;

                string sheetName = dimView.get_Parameter(BuiltInParameter.VIEWPORT_SHEET_NAME)?.AsString();
                if (string.IsNullOrEmpty(sheetName))
                    continue;

                result.Add(elem);
            }

            return result.ToArray();
        }

        private protected override CheckResultStatus Set_CheckerEntitiesHeap(Element[] elemColl)
        {
            _checkerEntitiesCollHeap.AddRange(CheckOverride(elemColl));
            _checkerEntitiesCollHeap.AddRange(CheckAccuracy(elemColl));

            return CheckResultStatus.Succeeded;
        }

        /// <summary>
        /// Определяю размеры, которые были переопределены в проекте
        /// </summary>
        private IEnumerable<CheckerEntity> CheckOverride(Element[] elemColl)
        {
            List<CheckerEntity> result = new List<CheckerEntity>();

            foreach (Element elem in elemColl)
            {
                if (!(elem is Dimension dim)) continue;

                // Игнорирую чертежные виды
                if (dim.View == null) continue;

                if (dim.View.GetType().Equals(typeof(ViewDrafting))) continue;

                CheckerEntity error = null;
                double? currentValue = dim.Value;
                if (currentValue.HasValue && dim.ValueOverride?.Length > 0)
                {
                    string overrideValue = dim.ValueOverride;
                    error = CheckDimValues(dim, currentValue.Value * 304.8, overrideValue);
                    if (error != null)
                        result.Add(error);
                }
                else
                {
                    DimensionSegmentArray segments = dim.Segments;
                    foreach (DimensionSegment segment in segments)
                    {
                        if (segment.ValueOverride?.Length > 0)
                        {
                            string overrideValue = segment.ValueOverride;
                            error = CheckDimValues(dim, segment.Value.Value * 304.8, overrideValue);
                            if (error != null)
                                result.Add(error);
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Анализ значения размера
        /// </summary>
        private CheckerEntity CheckDimValues(Dimension dim, double value, string overrideValue)
        {
            string[] splitValues = overrideValue.Split(_separArr, StringSplitOptions.None);
            
            // Анализирую диапозоны
            if (splitValues.Length > 1)
            {
                double overrideMinDouble = 0.0;
                for (int i = 0; i < splitValues.Length; i++)
                {
                    if (TryExtractNumber(splitValues[i], out overrideMinDouble)) 
                        break;
                }

                double overrideMaxDouble = 0.0;
                for (int j = splitValues.Length - 1; j >= 0; j--)
                {
                    if (TryExtractNumber(splitValues[j], out overrideMaxDouble))
                        break;
                }


                // Диапазон ОТ ... ДО
                if (overrideMinDouble != overrideMaxDouble)
                {
                    if (value > overrideMaxDouble || value < overrideMinDouble)
                        return new CheckerEntity(
                            dim,
                            "Нарушение диапозона (от ... до)",
                            "Размер вне диапозона",
                            $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а диапозон указан \"{overrideValue}\"");
                }
                // Одно значение с приставкой
                else
                {
                    string[] splitValues_MIN = overrideValue.Split(_separArr_MIN, StringSplitOptions.None);
                    string[] splitValues_MAX = overrideValue.Split(_separArr_MAX, StringSplitOptions.None);
                    if (splitValues_MIN.Length > 1 
                        && TryExtractNumber(splitValues_MIN[1], out overrideMinDouble)
                        && value < overrideMinDouble)
                    {
                        return new CheckerEntity(
                            dim,
                            "Нарушение диапозона (минимум)",
                            "Размер вне диапозона",
                            $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а диапозон указан \"{overrideValue}\"");
                    }
                    else if (splitValues_MAX.Length > 1
                        && TryExtractNumber(splitValues_MAX[1], out overrideMaxDouble)
                        && value > overrideMaxDouble)
                    {
                        return new CheckerEntity(
                            dim,
                            "Нарушение диапозона (максимум)",
                            "Размер вне диапозона",
                            $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а диапозон указан \"{overrideValue}\"");
                    }
                    else if (splitValues_MIN.Length < 2 && splitValues_MAX.Length < 2)
                        return new CheckerEntity(
                            dim,
                            "Нарушение диапозона",
                            "Не удалось определить данные. Нужен ручной анализ",
                            $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а диапозон указан \"{overrideValue}\"")
                            .Set_Status(ErrorStatus.Warning);
                }
            }

            // Нахожу значения без диапозона и игнорирую небольшие округления - больше 10 мм, при условии, что это не составляет 5% от размера
            else
            {
                if (!TryExtractNumber(overrideValue, out double overrideDouble))
                {
                    return new CheckerEntity(
                        dim,
                        "Нарушение переопределения размера",
                        "Не удалось определить данные. Нужен ручной анализ",
                        $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а при переопределении указано \"{overrideValue}\". Оцени вручную")
                        .Set_Status(ErrorStatus.Warning);
                }
                else if (Math.Abs(overrideDouble - value) > 10.0 || Math.Abs((overrideDouble / value) * 100 - 100) > 5)
                {
                    return new CheckerEntity(
                        dim,
                        "Нарушение переопределения размера",
                        "Размер значительно отличается от реального",
                        $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а при переопределении указано \"{overrideValue}\" мм. Разница существенная, лучше устранить.");
                }
                else
                {
                    return new CheckerEntity(
                        dim,
                        "Нарушение переопределения размера",
                        "Размер незначительно отличается от реального",
                        $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а при переопределении указано \"{overrideValue}\" мм. Разница не существенная, достаточно проконтролировать.")
                        .Set_Status(ErrorStatus.Warning);
                }
            }

            return null;
        }

        /// <summary>
        /// Определяю размеры, которые не соответствуют необходимой точности
        /// </summary>
        private IEnumerable<CheckerEntity> CheckAccuracy(Element[] elemColl)
        {
            List<CheckerEntity> result = new List<CheckerEntity>();

            HashSet<DimensionType> elemCollDimTypes = new HashSet<DimensionType>(
                elemColl.Select(elem => CheckDocument.GetElement(elem.GetTypeId())).OfType<DimensionType>(),
                new DimTypeCompare());

            foreach (DimensionType dimType in elemCollDimTypes)
            {
                DimensionStyleType dimStyleType = dimType.StyleType;
                if (dimStyleType == DimensionStyleType.Linear
                    || dimStyleType == DimensionStyleType.Diameter
                    || dimStyleType == DimensionStyleType.ArcLength
                    || dimStyleType == DimensionStyleType.LinearFixed)
                {
                    FormatOptions typeOpt = dimType.GetUnitsFormatOptions();
                    try
                    {
                        double currentAccuracy = typeOpt.Accuracy;
                        if (currentAccuracy > 1.0)
                        {
                            result.Add(new CheckerEntity(
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                                CheckDocument.GetElement(new ElementId(dimType.Id.IntegerValue)),
#else
                                CheckDocument.GetElement(new ElementId(dimType.Id.Value)),
#endif
                                "Нарушение точности в типе размера",
                                "Размер имеет запрещенно низкую точность",
                                $"Принятое округление в 1 мм, а в данном ТИПЕ - указано \"{currentAccuracy}\" мм. Замени округление, или удали типоразмер."));
                        }
                    }
                    catch (Exception)
                    {
                        //Игнорирую типы без настроек округления
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Извлекает первое число из строки. Поддерживает точку и запятую как 
        /// десятичный разделитель, неразрывные пробелы как разделители тысяч.
        /// </summary>
        private static bool TryExtractNumber(string input, out double result)
        {
            result = 0.0;
            if (string.IsNullOrWhiteSpace(input))
                return false;

            // Убираем разделители тысяч (обычные и неразрывные пробелы)
            string cleaned = input.Replace("\u00A0", "").Replace(" ", "");

            // Ищем первое число: опциональный знак + цифры + опциональная дробная часть
            Match m = Regex.Match(cleaned, @"[-+]?\d+(?:[.,]\d+)?");
            if (!m.Success)
                return false;

            // Нормализуем десятичный разделитель к точке
            string normalized = m.Value.Replace(',', '.');

            return double.TryParse(
                normalized,
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out result);
        }
    }
}
