using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib.Common;
using KPLN_ModelChecker_Lib.Common.ErrorTypes;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_Lib.ExternalCommands
{
    public sealed class CommandCheckDimensions : AbstractCheckCommand
    {
        /// <summary>
        /// Список сепараторов, для поиска диапозона у размеров
        /// </summary>
        private string[] _separArr = new string[]
        {
            "...",
            "до",
            "-",
            "max",
            "min"
        };

        public CommandCheckDimensions(UIApplication uiapp)
        {
            UIApp = uiapp;
            ErrorCollection = new List<ElementEntity>();
        }

        public override IList<ElementEntity> Run()
        {
            Document doc = UIApp.ActiveUIDocument.Document;
            FilteredElementCollector docDimensions = new FilteredElementCollector(doc).OfClass(typeof(Dimension)).WhereElementIsNotElementType();
            
            try
            {
                CheckOverride(doc, docDimensions);
                CheckAccuracy(doc);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            
            return ErrorCollection;
        }

        /// <summary>
        /// Определяю размеры, которые были переопределены в проекте
        /// </summary>
        private void CheckOverride(Document doc, FilteredElementCollector docDimensions)
        {
            foreach (Dimension dim in docDimensions)
            {
                // Игнорирую чертежные виды
                try
                {
                    if (dim.View.GetType().Equals(typeof(ViewDrafting))) { continue; }
                }
                catch (NullReferenceException) { continue; }

                double? currentValue = dim.Value;

                if (currentValue.HasValue && dim.ValueOverride?.Length > 0)
                {
                    double value = currentValue.Value * 304.8;
                    string overrideValue = dim.ValueOverride;
                    int dimId = dim.Id.IntegerValue;
                    string dimName = dim.Name;
                    CheckDimValues(doc, value, overrideValue, dimId, dimName);
                }
                else
                {
                    DimensionSegmentArray segments = dim.Segments;
                    foreach (DimensionSegment segment in segments)
                    {
                        if (segment.ValueOverride?.Length > 0)
                        {
                            double value = segment.Value.Value * 304.8;
                            string overrideValue = segment.ValueOverride;
                            int dimId = dim.Id.IntegerValue;
                            string dimName = dim.Name;
                            CheckDimValues(doc, value, overrideValue, dimId, dimName);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Анализ значения размера
        /// </summary>
        private void CheckDimValues(Document doc, double value, string overrideValue, int elemId, string elemName)
        {

            string ovverrideMinValue = String.Empty;
            double overrideMinDouble = 0.0;
            string ovverrideMaxValue = String.Empty;
            double overrideMaxDouble = 0.0;

            string[] splitValues = overrideValue.Split(_separArr, StringSplitOptions.None);

            // Анализирую диапозоны
            if (splitValues.Length > 1)
            {
                ovverrideMinValue = splitValues[0];
                if (ovverrideMinValue.Length == 0)
                {
                    ovverrideMinValue = overrideValue;
                }
                else
                {
                    ovverrideMaxValue = splitValues[1];
                }

                string onlyNumbMin = new string(ovverrideMinValue.Where(x => Char.IsDigit(x)).ToArray());
                Double.TryParse(onlyNumbMin, out overrideMinDouble);
                if (!ovverrideMaxValue.Equals(String.Empty))
                {
                    string onlyNumbMax = new string(ovverrideMaxValue.Where(x => Char.IsDigit(x)).ToArray());
                    Double.TryParse(onlyNumbMax, out overrideMaxDouble);
                    // Нахожу значения вне диапозоне
                    if (value >= overrideMaxDouble | value < overrideMinDouble)
                    {
                        ErrorCollection.Add(new ElementEntity(
                            doc.GetElement(new ElementId(elemId)),
                            elemName,
                            "Размер вне диапозона",
                            $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а диапозон указан \"{overrideValue}\"",
                            new Error())
                        );
                    }
                }
                else
                {
                    ErrorCollection.Add(new ElementEntity(
                        doc.GetElement(new ElementId(elemId)),
                        elemName,
                        "Диапазон. Нужен ручной анализ",
                        $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а диапозон указан \"{overrideValue}\"",
                        new Warning())
                    );
                }
            }

            // Нахожу значения без диапозона и игнорирую небольшие округления - больше 10 мм, при условии, что это не составляет 5% от размера
            else
            {
                double overrideDouble = 0.0;
                string onlyNumbMin = new string(overrideValue.Where(x => Char.IsDigit(x)).ToArray());
                Double.TryParse(onlyNumbMin, out overrideDouble);
                if (overrideDouble == 0.0)
                {
                    ErrorCollection.Add(new ElementEntity(
                        doc.GetElement(new ElementId(elemId)),
                        elemName,
                        "Нет возможности анализа",
                        $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а при переопределении указано \"{overrideValue}\". Оцени вручную",
                        new Error())
                    );
                }
                else if (Math.Abs(overrideDouble - value) > 10.0 || Math.Abs((overrideDouble / value) * 100 - 100) > 5)
                {
                    ErrorCollection.Add(new ElementEntity(
                        doc.GetElement(new ElementId(elemId)),
                        elemName,
                        "Размер значительно отличается от реального",
                        $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а при переопределении указано \"{overrideValue}\" мм. Разница существенная, лучше устранить.",
                        new Error())
                    );
                }
                else
                {
                    ErrorCollection.Add(new ElementEntity(
                        doc.GetElement(new ElementId(elemId)),
                        elemName,
                        "Размер не значительно отличается от реального",
                        $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а при переопределении указано \"{overrideValue}\" мм. Разница не существенная, достаточно проконтролировать.",
                        new LittleWarning())
                    );
                }
            }
        }

        /// <summary>
        /// Определяю размеры, которые не соответствуют необходимой точности
        /// </summary>
        private void CheckAccuracy(Document doc)
        {
            FilteredElementCollector docDimensionTypes = new FilteredElementCollector(doc).OfClass(typeof(DimensionType)).WhereElementIsElementType();
            foreach (DimensionType dimType in docDimensionTypes)
            {
                if (dimType.UnitType == UnitType.UT_Length)
                {
                    FormatOptions typeOpt = dimType.GetUnitsFormatOptions();
                    try
                    {
                        double currentAccuracy = typeOpt.Accuracy;
                        if (currentAccuracy > 1.0)
                        {
                            ErrorCollection.Add(new ElementEntity(
                                doc.GetElement(new ElementId(dimType.Id.IntegerValue)),
                                dimType.Name,
                                "Размер имеет запрещенно низкую точность",
                                $"Принятое округление в 1 мм, а в данном типе - указано \"{currentAccuracy}\" мм. Замени округление, или удали типоразмер.",
                                new Error())
                            );
                        }
                    }
                    catch (Exception)
                    {
                        //Игнорирую типы без настроек округления
                    }
                }
            }
        }
    }
}
