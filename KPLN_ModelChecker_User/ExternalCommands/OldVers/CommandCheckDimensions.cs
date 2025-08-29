using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_PluginActivityWorker;
using KPLN_ModelChecker_Lib;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    internal class DimTypeCompare : IEqualityComparer<DimensionType>
    {
        public bool Equals(DimensionType x, DimensionType y) => x.Id.Equals(y.Id);

        public int GetHashCode(DimensionType obj) => obj.Id.GetHashCode();
    }


    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandCheckDimensions : AbstrCheckCommandOld<CommandCheckDimensions>, IExternalCommand
    {
        internal const string PluginName = "Проверка размеров";

        /// <summary>
        /// Список сепараторов, для поиска диапозона у размеров
        /// </summary>
        private readonly string[] _separArr = new string[]
        {
            "...",
            "до",
            "-",
            "max",
            "макс",
            "min",
            "мин"
        };

        public CommandCheckDimensions() : base()
        {
        }

        internal CommandCheckDimensions(ExtensibleStorageEntity esEntity) : base(esEntity)
        {
        }

        /// <summary>
        /// Реализация IExternalCommand
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return ExecuteByUIApp(commandData.Application);
        }

        public override Result ExecuteByUIApp(UIApplication uiapp)
        {
            DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName($"{PluginName}", ModuleData.ModuleName).ConfigureAwait(false);

            _uiApp = uiapp;

            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            string docTitle = doc.Title;


            // Собираю листы модели
            ViewSheet[] docSheetColl;
            IEnumerable<ViewSheet> docAllList = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Sheets)
                .WhereElementIsNotElementType()
                .Cast<ViewSheet>();
            // Кастомно для проектов на основе моделей субчиков (например только нужную стадию)
            if (docTitle.Contains("СЕТ_1") && (docTitle.Contains("_КЖ") || docTitle.Contains("_КМ")))
                docSheetColl = docAllList
                    .Where(vsh => vsh.LookupParameter("Орг.КомплектЧертежей")?.AsString()?.ToLower().Contains("кж") == true)
                    .ToArray();
            // Стандартно - для всех листов
            else
                docSheetColl = docAllList.ToArray();


            // Собираю размеры с видов на листах
            List<Element> dimToCheck = new List<Element>();
            List<Element> docDimensions = new FilteredElementCollector(doc)
                .OfClass(typeof(Dimension))
                .WhereElementIsNotElementType()
                .ToList();
            foreach (Element elem in docDimensions)
            {
                Dimension dim = (Dimension)elem;
                View dimView = dim.View;
                // Такое может быть, т.к. этого же класса зависимости (сам размер может быть удален)
                if (dimView == null)
                    continue;

                // Если листа нет, то анализировать нечего. Коллекция листов готовиться заранее (по нужному парамтеру)
                string sheetName = dimView.get_Parameter(BuiltInParameter.VIEWPORT_SHEET_NAME)?.AsString();
                if (string.IsNullOrEmpty(sheetName))
                    continue;

                dimToCheck.Add(elem);
            }


            #region Проверяю и обрабатываю элементы
            WPFEntity[] wpfColl = CheckCommandRunner(doc, dimToCheck.ToArray());
            OutputMainForm form = ReportCreatorAndDemonstrator(doc, wpfColl);
            if (form != null) form.Show();
            else return Result.Cancelled;
            #endregion

            return Result.Succeeded;
        }

        private protected override IEnumerable<CheckCommandError> CheckElements(Document doc, object[] objColl) => Enumerable.Empty<CheckCommandError>();

        private protected override IEnumerable<WPFEntity> PreapareElements(Document doc, Element[] elemColl)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            result.AddRange(CheckOverride(elemColl));
            result.AddRange(CheckAccuracy(doc, elemColl));

            return result;
        }

        private protected override void SetWPFEntityFiltration(WPFReportCreator report)
        {
            report.SetWPFEntityFiltration_ByErrorHeader();
        }

        /// <summary>
        /// Определяю размеры, которые были переопределены в проекте
        /// </summary>
        private IEnumerable<WPFEntity> CheckOverride(Element[] elemColl)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            foreach (Element elem in elemColl)
            {
                if (!(elem is Dimension dim)) continue;

                // Игнорирую чертежные виды
                if (dim.View == null) continue;

                if (dim.View.GetType().Equals(typeof(ViewDrafting))) continue;

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
                            ESEntity,
                            dim,
                            "Нарушение диапозона",
                            "Размер вне диапозона",
                            $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а диапозон указан \"{overrideValue}\"",
                            false);
                    }
                }
                else
                {
                    return new WPFEntity(
                        ESEntity,
                        dim,
                        "Нарушение диапозона",
                        "Не удалось определить данные. Нужен ручной анализ",
                        $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а диапозон указан \"{overrideValue}\"",
                        false,
                        ErrorStatus.Warning);
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
                        ESEntity,
                        dim,
                        "Нарушение переопределения размера",
                        "Не удалось определить данные. Нужен ручной анализ",
                        $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а при переопределении указано \"{overrideValue}\". Оцени вручную",
                        false,
                        ErrorStatus.Warning);
                }
                else if (Math.Abs(overrideDouble - value) > 10.0 || Math.Abs((overrideDouble / value) * 100 - 100) > 5)
                {
                    return new WPFEntity(
                        ESEntity,
                        dim,
                        "Нарушение переопределения размера",
                        "Размер значительно отличается от реального",
                        $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а при переопределении указано \"{overrideValue}\" мм. Разница существенная, лучше устранить.",
                        false);
                }
                else
                {
                    return new WPFEntity(
                        ESEntity,
                        dim,
                        "Нарушение переопределения размера",
                        "Размер незначительно отличается от реального",
                        $"Значение реального размера \"{Math.Round(value, 2)}\" мм, а при переопределении указано \"{overrideValue}\" мм. Разница не существенная, достаточно проконтролировать.",
                        false,
                        ErrorStatus.Warning);
                }
            }

            return null;
        }

        /// <summary>
        /// Определяю размеры, которые не соответствуют необходимой точности
        /// </summary>
        private IEnumerable<WPFEntity> CheckAccuracy(Document doc, Element[] elemColl)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            HashSet<DimensionType> elemCollDimTypes = new HashSet<DimensionType>(
                elemColl.Select(elem => doc.GetElement(elem.GetTypeId())).Cast<DimensionType>(),
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
                            result.Add(new WPFEntity(
                                ESEntity,
                                doc.GetElement(new ElementId(dimType.Id.IntegerValue)),
                                "Нарушение точности в типе размера",
                                "Размер имеет запрещенно низкую точность",
                                $"Принятое округление в 1 мм, а в данном ТИПЕ - указано \"{currentAccuracy}\" мм. Замени округление, или удали типоразмер.",
                                false));
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
