using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_ModelChecker_User.Common.Collections;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandCheckDimensions : AbstrCheckCommand, IExternalCommand
    {
        /// <summary>
        /// Список сепараторов, для поиска диапозона у размеров
        /// </summary>
        private readonly string[] _separArr = new string[]
        {
            "...",
            "до",
            "-",
            "max",
            "min"
        };

        internal static CommandCheckDimensions()
        {

        }

        /// <summary>
        /// Реализация IExternalCommand
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return Execute(commandData.Application);
        }

        internal override Result Execute(UIApplication uiapp)
        {
            CheckName = "Проверка размеров";
            MainStorageName = "KPLN_CheckDimensions";
            LastRunGuid = new Guid("f2e615e0-a15b-43df-a199-a88d18a2f568");
            UserTextGuid = new Guid("f2e615e0-a15b-43df-a199-a88d18a2f569");
            
            _application = uiapp;

            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Получаю коллекцию элементов для анализа
            Element[] docDimensions = new FilteredElementCollector(doc).OfClass(typeof(Dimension)).WhereElementIsNotElementType().ToArray();

            #region Проверяю и обрабатываю элементы
            WPFEntity[] wpfColl = CheckCommandRunner(doc, docDimensions);
            OutputMainForm form = ReportCreatorAndDemonstrator(doc, wpfColl);
            if (form != null) form.Show();
            else return Result.Cancelled;
            #endregion

            return Result.Succeeded;
        }

        private protected override IEnumerable<CheckCommandError> CheckElements(Document doc, Element[] elemColl) => Enumerable.Empty<CheckCommandError>();

        private protected override IEnumerable<WPFEntity> PreapareElements(Document doc, Element[] elemColl)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            result.AddRange(CheckOverride(doc, elemColl));
            result.AddRange(CheckAccuracy(doc));

            return result;
        }

        private protected override void SetWPFEntityFiltration(WPFReportCreator report)
        {
            report.SetWPFEntityFiltration_ByErrorHeader();
        }

        /// <summary>
        /// Определяю размеры, которые были переопределены в проекте
        /// </summary>
        private IEnumerable<WPFEntity> CheckOverride(Document doc, Element[] elemColl)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            foreach (Dimension dim in elemColl)
            {
                // Игнорирую чертежные виды
                if (dim.View == null)
                    continue;

                if (dim.View.GetType().Equals(typeof(ViewDrafting)))
                    continue;

                WPFEntity error = null;
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
        private WPFEntity CheckDimValues(Dimension dim, double value, string overrideValue)
        {
            int dimId = dim.Id.IntegerValue;
            string dimName = dim.Name;

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
                    ovverrideMinValue = overrideValue;
                else
                    ovverrideMaxValue = splitValues[1];

                string onlyNumbMin = new string(ovverrideMinValue.Where(x => Char.IsDigit(x)).ToArray());
                Double.TryParse(onlyNumbMin, out overrideMinDouble);
                if (!ovverrideMaxValue.Equals(String.Empty))
                {
                    string onlyNumbMax = new string(ovverrideMaxValue.Where(x => Char.IsDigit(x)).ToArray());
                    Double.TryParse(onlyNumbMax, out overrideMaxDouble);
                    // Нахожу значения вне диапозоне
                    if (value >= overrideMaxDouble | value < overrideMinDouble)
                    {
                        return new WPFEntity(
                            dim,
                            Status.Error,
                            "Нарушение диапозона",
                            "Размер вне диапозона",
                            false,
                            false,
                            $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а диапозон указан \"{overrideValue}\"");
                    }
                }
                else
                {
                    return new WPFEntity(
                        dim,
                        Status.Warning,
                        "Нарушение диапозона",
                        "Не удалось определить данные. Нужен ручной анализ",
                        false,
                        false,
                        $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а диапозон указан \"{overrideValue}\"");
                }
            }

            // Нахожу значения без диапозона и игнорирую небольшие округления - больше 10 мм, при условии, что это не составляет 5% от размера
            else
            {
                string onlyNumbMin = new string(overrideValue.Where(x => Char.IsDigit(x)).ToArray());
                Double.TryParse(onlyNumbMin, out double overrideDouble);
                if (overrideDouble == 0.0)
                {
                    return new WPFEntity(
                        dim,
                        Status.Warning,
                        "Нарушение переопределения размера",
                        "Не удалось определить данные. Нужен ручной анализ",
                        false,
                        false,
                        $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а при переопределении указано \"{overrideValue}\". Оцени вручную");
                }
                else if (Math.Abs(overrideDouble - value) > 10.0 || Math.Abs((overrideDouble / value) * 100 - 100) > 5)
                {
                    return new WPFEntity(
                        dim,
                        Status.Error,
                        "Нарушение переопределения размера",
                        "Размер значительно отличается от реального",
                        false,
                        false,
                        $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а при переопределении указано \"{overrideValue}\" мм. Разница существенная, лучше устранить.");
                }
                else
                {
                    return new WPFEntity(
                        dim,
                        Status.Warning,
                        "Нарушение переопределения размера",
                        "Размер незначительно отличается от реального",
                        false,
                        false,
                        $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а при переопределении указано \"{overrideValue}\" мм. Разница не существенная, достаточно проконтролировать.");
                }
            }

            return null;
        }

        /// <summary>
        /// Определяю размеры, которые не соответствуют необходимой точности
        /// </summary>
        private IEnumerable<WPFEntity> CheckAccuracy(Document doc)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            FilteredElementCollector docDimensionTypes = new FilteredElementCollector(doc).OfClass(typeof(DimensionType)).WhereElementIsElementType();
            foreach (DimensionType dimType in docDimensionTypes)
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
                            result.Add(new WPFEntity(
                                doc.GetElement(new ElementId(dimType.Id.IntegerValue)),
                                Status.Error,
                                "Нарушение точности размера",
                                "Размер имеет запрещенно низкую точность",
                                false,
                                false,
                                $"Принятое округление в 1 мм, а в данном типе - указано \"{currentAccuracy}\" мм. Замени округление, или удали типоразмер."));
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
    }
}
