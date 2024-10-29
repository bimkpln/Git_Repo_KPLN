using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text.RegularExpressions;

namespace KPLN_ModelChecker_Lib.LevelAndGridBoxUtil.Common
{
    [SuppressMessage("ReSharper", "ForeachCanBeConvertedToQueryUsingAnotherGetEnumerator")]
    public class LevelData
    {
        internal static readonly string ParLvlName = "ПАР";
        internal static readonly string StilLvlName = "СТЛ";
        internal static readonly string SectLvlName = "С";
        internal static readonly string KorpLvlName = "К";

        /// <summary>
        /// Номер секции для проектов, котоыре не деляться на секции (т.е. она одна)
        /// </summary>
        private static string _singleSectionNumber;
        private double[] _minAndMaxLvlPnts;

        /// <summary>
        /// Текущий уровень
        /// </summary>
        public Level CurrentLevel { get; private set; }

        /// <summary>
        /// Уровень выше для секции
        /// </summary>
        public Level CurrentAboveLevel { get; private set; }

        /// <summary>
        /// Уровень ниже для секции
        /// </summary>
        public Level CurrentDownLevel { get; private set; }

        /// <summary>
        /// Нижняя и верхняя точки для формирования бокса между уровнями
        /// </summary>
        public double[] MinAndMaxLvlPnts
        {
            get
            {
                if (_minAndMaxLvlPnts == null)
                {
                    _minAndMaxLvlPnts = new double[2];

                    double minPointOfLevels;
                    if (CurrentDownLevel == null)
                        minPointOfLevels = CurrentLevel.Elevation - DownAndTopExtra - FloorScreedHeight;
                    else
                        minPointOfLevels = CurrentLevel.Elevation - FloorScreedHeight;

                    double maxPointOLevels;
                    if (CurrentAboveLevel == null)
                        maxPointOLevels = minPointOfLevels + DownAndTopExtra - FloorScreedHeight;
                    else
                        maxPointOLevels = CurrentAboveLevel.Elevation - FloorScreedHeight;

                    _minAndMaxLvlPnts[0] = minPointOfLevels;
                    _minAndMaxLvlPnts[1] = maxPointOLevels;
                }

                return _minAndMaxLvlPnts;
            }
        }

        /// <summary>
        /// Толщина смещения относительно уровня (чаще всего - стяжка пола). Нужна для перекидки значения элементов в стяжке на этаж выше
        /// </summary>
        public double FloorScreedHeight { get; }

        /// <summary>
        /// Размер увеличения нижнего и верхнего боксов. Нужна для привязки элементов, расположенных за пределами крайних уровней
        /// </summary>
        public double DownAndTopExtra { get; }

        /// <summary>
        /// Номер уровня
        /// </summary>
        public string CurrentLevelNumber { get; private set; }

        /// <summary>
        /// Секции для уровня
        /// </summary>
        public string CurrentSectionNumber { get; private set; }

        private LevelData(Level level, string lvlNumber, string sectNumber, Level aboveLevel, Level downLevel, double floorScreedHeight, double downAndTopExtra)
        {
            CurrentLevel = level;
            CurrentLevelNumber = lvlNumber;
            CurrentSectionNumber = sectNumber;
            CurrentAboveLevel = aboveLevel;
            CurrentDownLevel = downLevel;
            FloorScreedHeight = floorScreedHeight;
            DownAndTopExtra = downAndTopExtra;
        }

        /// <summary>
        /// Подготовка коллекции уровней для анализа
        /// </summary>
        /// <param name="doc">Revit-документ для анализа</param>
        /// <param name="floorScreedHeight">Толщина стяжки пола АР</param>
        /// <param name="downAndTopExtra">Расширение дампозона для самого нижнего и самого верхнего уровней</param>
        /// <param name="sectSeparParamName">Параметр для разделения осей и уровней по секциям (ОБЯЗАТЕЛЬНО заполнять разделителями как по ВЕР для блока "Секция/Корпус")</param>
        /// <param name="levelIndexParamName">Параметр для разделения уровней по этажам</param>
        /// <param name="multiGridsSet">Массив из одного номер секции, если здание с одной секцией</param>
        [SuppressMessage("ReSharper", "LoopCanBeConvertedToQuery")]
        internal static List<LevelData> LevelPrepare(
            Document doc, 
            double floorScreedHeight, 
            double downAndTopExtra, 
            string sectSeparParamName, 
            string levelIndexParamName, 
            HashSet<string> multiGridsSet = null)
        {
            if (multiGridsSet != null && multiGridsSet.Count() != 1)
                throw new CheckerException(
                    $"Системная ошибка - отправь разработчику: нарушена логика для односекционных зданий");
            
            if (multiGridsSet != null)
                _singleSectionNumber = multiGridsSet.FirstOrDefault();
            
            List<LevelData> prepareLevels = new List<LevelData>();


            Level[] levelColl;
            if (doc.Title.StartsWith("СЕТ_1"))
            {
                levelColl = new FilteredElementCollector(doc)
                    .OfClass(typeof(Level))
                    .WhereElementIsNotElementType()
                    .Cast<Level>()
                    .OrderBy(x => x.Elevation)
                    .Where(x => !x.Name.Contains("КР_"))
                    .ToArray();
            }
            else
            {
                levelColl = new FilteredElementCollector(doc)
                    .OfClass(typeof(Level))
                    .WhereElementIsNotElementType()
                    .Cast<Level>()
                    .OrderBy(x => x.Elevation)
                    .ToArray();
            }

            foreach (Level level in levelColl)
            {
                string lvlNumber = GetLevelNumber(level, levelIndexParamName);
                List<string> lvlSections = GetLevelSections(level, sectSeparParamName);
                foreach (string levelSection in lvlSections)
                {
                    Level aboveLevel = GetAboveLevel(levelColl, level, levelSection, levelIndexParamName,
                        sectSeparParamName);
                    Level downLevel = GetDownLevel(levelColl, level, levelSection, levelIndexParamName,
                        sectSeparParamName);
                    LevelData myLevel = new LevelData(level, lvlNumber, levelSection, aboveLevel, downLevel,
                        floorScreedHeight, downAndTopExtra);
                    
                    prepareLevels.Add(myLevel);
                }
            }

            return prepareLevels;
        }

        /// <summary>
        /// Взять индекс уровня (номер этажа)
        /// </summary>
        /// <param name="level">Уровень на проверку</param>
        /// <param name="levelIndexParamName">Параметр для разделения уровней по этажам</param>
        public static string GetLevelNumber(Level level, string levelIndexParamName)
        {
            // Блок, когда параметр индекса уровня заполнен
            Parameter lvlIndParam = level.LookupParameter(levelIndexParamName);
            if (lvlIndParam != null && !string.IsNullOrEmpty(lvlIndParam.AsString()))
            {
                string result = lvlIndParam.AsString();
                // Подрезаю имена для СЕТ_1
                if (result.Contains("_этаж"))
                    result = result.Split(new string[] { "_этаж" }, StringSplitOptions.None)[0];
                if (result.Contains("_Кровля"))
                    result = result.Split(new string[] { "_Кровля" }, StringSplitOptions.None)[0];
                
                return result;
            }
            // Блок, когда индекс пытаемся найти по имени уровня (по ВЕР КПЛН)
            else
            {
                string levname = level.Name;
                string[] splitname = levname.Split('_');
                if (splitname.Length < 2)
                    throw new CheckerException(
                        $"Некорректное имя уровня (см. ВЕР, паттерн: С1_01_+0.000_Технический этаж). Id: {level.Id}");

                // Обработка проектов без деления на секции
                return splitname[0].Any(char.IsLetter) ? splitname[1] : splitname[0];
            }
        }

        /// <summary>
        /// Взять секции, к которым относится уровень
        /// </summary>
        /// <param name="level">Уровень на проверку</param>
        /// <param name="sectSeparParamName">Параметр для разделения осей и уровней по секциям
        /// (ОБЯЗАТЕЛЬНО заполнять разделителями как по ВЕР для блока "Секция/Корпус")</param>
        private static List<string> GetLevelSections(Level level, string sectSeparParamName)
        {
            // Блок, когда параметр секции уровня заполнен
            Parameter sectParam = level.LookupParameter(sectSeparParamName);
            if (sectParam != null && !string.IsNullOrEmpty(sectParam.AsString()))
                return GetSections(sectParam.AsString());
            // Блок, когда секцию пытаемся найти по имени уровня (по ВЕР КПЛН)
            else
            {
                string levname = level.Name;
                string[] splitname = levname.Split('_');
                if (splitname.Length < 2)
                    throw new CheckerException(
                        $"Некорректное имя уровня (см. ВЕР, паттерн: С1_01_+0.000_Технический этаж). Id: {level.Id}");

                if (splitname[0].Any(char.IsLetter)
                    && !splitname[0].Contains(SectLvlName) 
                    && !splitname[0].Contains(KorpLvlName) 
                    && !splitname[0].Contains(ParLvlName) 
                    && !splitname[0].Contains(StilLvlName))
                    throw new CheckerException(
                        $"Некорректное имя уровня (см. ВЕР, кодировка секции/корпуса: С1/С1.2/С1-6/К1/ПАР/СТЛ. Id: {level.Id}");

                // Обработка проектов без деления на секции
                if (!splitname[0].Any(char.IsLetter)
                    && !string.IsNullOrEmpty(_singleSectionNumber))
                    return new List<string> { _singleSectionNumber };
            
                return GetSections(splitname[0]);
            }
        }

        private static List<string> GetSections(string input)
        {
            List<string> result = new List<string>();
            
            // Это для корпусов-одиночек.
            if (string.Equals(input, ParLvlName, StringComparison.OrdinalIgnoreCase))
                result.Add(ParLvlName);
            else if (string.Equals(input, StilLvlName, StringComparison.OrdinalIgnoreCase))
                result.Add(StilLvlName);
            else if (string.Equals(input, KorpLvlName, StringComparison.OrdinalIgnoreCase))
                result.Add(KorpLvlName);
            else if (string.Equals(input, SectLvlName, StringComparison.OrdinalIgnoreCase))
                result.Add(SectLvlName);
            // Это для сложных структур
            else
            {
                string[] splitParts = input.Split('/');
                // Для заполнения через параметр в формате:
                // "К1С1-С4/К2С1-С3/К3-К4С3-С4" => "К1С1, К1С2, К1С3, К1С4, К2С1, К2С2, К2С3, К3С3, К3С4, К4С3, К4С4"
                if (splitParts.Length > 1)
                {
                    foreach (string part in splitParts)
                    {
                        Match match = Regex.Match(part, @"К(\d+)(?:-К(\d+))?С(\d+)(?:-С(\d+))?");
                        // Если нашли диапазоны для К и/или С
                        if (match.Success)
                        {
                            int kStart = int.Parse(match.Groups[1].Value);
                            int kEnd = match.Groups[2].Success ? int.Parse(match.Groups[2].Value) : kStart;
                            int sStart = int.Parse(match.Groups[3].Value);
                            int sEnd = match.Groups[4].Success ? int.Parse(match.Groups[4].Value) : sStart;

                            for (int k = kStart; k <= kEnd; k++)
                            {
                                for (int s = sStart; s <= sEnd; s++)
                                {
                                    result.Add($"К{k}С{s}");
                                }
                            }
                        }
                        // Если диапазона нет, добавляем элемент как есть
                        else
                            result.Add(part);
                    }
                }
                // Для анализа имени уровня по ВЕР
                else
                {
                    string resultBuildSepar = input.Contains(KorpLvlName) ? KorpLvlName : SectLvlName;
                    MatchCollection matches = Regex.Matches(input, @"(\d+)-?(\d*)");
                    foreach (Match match in matches)
                    {
                        int startRange = int.Parse(match.Groups[1].Value);
                        int endRange = string.IsNullOrEmpty(match.Groups[2].Value) ? startRange : int.Parse(match.Groups[2].Value);
                        // Добавляем числа в указанном диапазоне в результат
                        for (int i = startRange; i <= endRange; i++)
                        {
                            result.Add($"{resultBuildSepar}{i}");
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Взять уровень на 1 выше по индексу этажа и номеру секции
        /// </summary>
        /// <param name="levelColl">Коллекция уровней для проверки</param>
        /// <param name="checkLevel">Уровень, по которому проверяем</param>
        /// <param name="levelSection">Номер секции</param>
        /// <param name="levelIndexParamName">Параметр для разделения уровней по этажам</param>
        /// <param name="sectSeparParamName">Параметр для разделения осей и уровней по секциям</param>
        private static Level GetAboveLevel(Level[] levelColl, Level checkLevel, string levelSection,
            string levelIndexParamName, string sectSeparParamName)
        {
            string chkLvlNumber = GetLevelNumber(checkLevel, levelIndexParamName);
            if (int.TryParse(chkLvlNumber, out int chkNumber))
            {
                if (chkNumber == 0)
                    throw new CheckerException($"У уровня id {checkLevel.Id} значение уровня - 0. Это запрещено");

                foreach (Level level in levelColl)
                {
                    string lvlNumber = GetLevelNumber(level, levelIndexParamName);
                    var a = GetLevelSections(level, sectSeparParamName);
                    if (chkLvlNumber.Equals(lvlNumber)
                        || !GetLevelSections(level, sectSeparParamName).Contains(levelSection))
                        continue;

                    if (int.TryParse(lvlNumber, out int number))
                    {
                        if (number == 0)
                            throw new CheckerException($"У уровня id {level.Id} значение уровня - 0. Это запрещено");

                        if (chkNumber < 0
                            && number > 0
                            && chkNumber - number == -2)
                            return level;
                        else if (number - chkNumber == 1)
                            return level;
                    }
                    else
                        throw new CheckerException($"У уровня с id {level.Id} не удалось получить int его номера. Обратись к разработчику");
                }
            }
            else
                throw new CheckerException(
                    $"У уровня с id {checkLevel.Id} не удалось получить int его номера. Варианты ошибок:" +
                    $"\n 1. Здание многосекционное/многокорпусное, но у уровня секции не указаны." +
                    $"\n 2. Данные из параметра секции не совпадают с блоком в имени уровня, который отвечает за секцию/корпус");

            return null;
        }

        /// <summary>
        /// Взять уровень на 1 ниже по индексу этажа и номеру секции
        /// </summary>
        /// <param name="levelColl">Коллекция уровней для проверки</param>
        /// <param name="checkLevel">Уровень, по которому проверяем</param>
        /// <param name="levelSection">Номер секции</param>
        /// <param name="levelIndexParamName">Параметр для разделения уровней по этажам</param>
        /// <param name="sectSeparParamName">Параметр для разделения осей и уровней по секциям</param>
        private static Level GetDownLevel(Level[] levelColl, Level checkLevel, string levelSection,
            string levelIndexParamName, string sectSeparParamName)
        {
            string chkLvlNumber = GetLevelNumber(checkLevel, levelIndexParamName);
            if (int.TryParse(chkLvlNumber, out int chkNumber))
            {
                if (chkNumber == 0)
                    throw new CheckerException($"У уровня id {checkLevel.Id} значение уровня - 0. Это запрещено");

                foreach (Level level in levelColl)
                {
                    string lvlNumber = GetLevelNumber(level, levelIndexParamName);
                    if (chkLvlNumber.Equals(lvlNumber)
                        || !GetLevelSections(level, sectSeparParamName).Contains(levelSection))
                        continue;

                    if (int.TryParse(lvlNumber, out int number))
                    {
                        if (number == 0)
                            throw new CheckerException($"У уровня id {level.Id} значение уровня - 0. Это запрещено");

                        if (chkNumber > 0
                            && number < 0
                            && chkNumber - number == 2
                            )
                            return level;
                        else if (chkNumber - number == 1)
                            return level;
                    }
                    else
                        throw new CheckerException($"У уровня id {level.Id} не удалось получить int его номера");
                }
            }
            else
                throw new CheckerException($"У уровня id {checkLevel.Id} не удалось получить int его номера");

            return null;
        }
    }
}
