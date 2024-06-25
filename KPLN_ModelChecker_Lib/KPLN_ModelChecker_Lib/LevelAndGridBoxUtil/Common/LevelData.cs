using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace KPLN_ModelChecker_Lib.LevelAndGridBoxUtil.Common
{
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
        /// Размер увеличения нижнего и вехнего боксов. Нужна для привязки элементов, расположенных за пределами крайних уровней
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
        /// <param name="floorScreedHeight">Толщина стяжки пола АР</param>
        /// <param name="singleSectionNumber">Номер секции, если здание с одной секцией</param>
        internal static List<LevelData> LevelPrepare(Document doc, double floorScreedHeight, double downAndTopExtra, string singleSectionNumber = null)
        {
            _singleSectionNumber = singleSectionNumber;
            List<LevelData> preapareLevels = new List<LevelData>();

            Level[] levelColl = new FilteredElementCollector(doc)
                    .OfClass(typeof(Level))
                    .WhereElementIsNotElementType()
                    .Cast<Level>()
                    .OrderBy(x => x.Elevation)
                    .ToArray();

            foreach (Level level in levelColl)
            {
                string lvlNumber = GetLevelNumber(level);
                List<string> lvlSections = GetLevelSections(level);
                foreach (string levelSection in lvlSections)
                {
                    Level aboveLevel = GetAboveLevel(levelColl, level, levelSection);
                    Level downLevel = GetDownLevel(levelColl, level, levelSection);
                    LevelData myLevel = new LevelData(level, lvlNumber, levelSection, aboveLevel, downLevel, floorScreedHeight, downAndTopExtra);
                    preapareLevels.Add(myLevel);
                }
            }

            return preapareLevels;
        }

        /// <summary>
        /// Взять индекс уровня (номер этажа)
        /// </summary>
        /// <param name="level">Уровень на проверку</param>
        public static string GetLevelNumber(Level level)
        {
            string levname = level.Name;
            string[] splitname = levname.Split('_');
            if (splitname.Length < 2)
                throw new CheckerException($"Некорректное имя уровня (см. ВЕР, паттерн: С1_01_+0.000_Технический этаж). Id: {level.Id}");

            // Обработка проектов без деления на секции
            if(!splitname[0].Any(char.IsLetter)
                && !string.IsNullOrEmpty(_singleSectionNumber))
                return splitname[0];
            else
                return splitname[1];
        }

        /// <summary>
        /// Взять секции, к которым относится уровень
        /// </summary>
        /// <param name="level"></param>
        private static List<string> GetLevelSections(Level level)
        {
            string levname = level.Name;
            string[] splitname = levname.Split('_');
            if (splitname.Length < 2)
                throw new CheckerException($"Некорректное имя уровня (см. ВЕР, паттерн: С1_01_+0.000_Технический этаж). Id: {level.Id}");

            if (splitname[0].Any(char.IsLetter)
                && !splitname[0].Contains(SectLvlName) 
                && !splitname[0].Contains(KorpLvlName) 
                && !splitname[0].Contains(ParLvlName) 
                && !splitname[0].Contains(StilLvlName))
                throw new CheckerException($"Некорректное имя уровня (см. ВЕР, кодировка секции/корпуса: С1/С1.2/С1-6/К1/ПАР/СТЛ. Id: {level.Id}");

            // Обработка проектов без деления на секции
            if (!splitname[0].Any(char.IsLetter)
                && !string.IsNullOrEmpty(_singleSectionNumber))
                return new List<string> { _singleSectionNumber };
            
            return GetSections(splitname[0]);
        }

        private static List<string> GetSections(string input)
        {
            List<string> result = new List<string>();

            if (string.Equals(input, ParLvlName, StringComparison.OrdinalIgnoreCase))
                result.Add(ParLvlName);
            else if(string.Equals(input, StilLvlName, StringComparison.OrdinalIgnoreCase))
                result.Add(StilLvlName);
            else if (string.Equals(input, KorpLvlName, StringComparison.OrdinalIgnoreCase))
                result.Add(KorpLvlName);
            else if (string.Equals(input, SectLvlName, StringComparison.OrdinalIgnoreCase))
                result.Add(SectLvlName);

            MatchCollection matches = Regex.Matches(input, @"(\d+)-?(\d*)");
            foreach (Match match in matches)
            {
                int startRange = int.Parse(match.Groups[1].Value);
                int endRange = string.IsNullOrEmpty(match.Groups[2].Value) ? startRange : int.Parse(match.Groups[2].Value);
                // Добавляем числа в указанном диапазоне в результат
                for (int i = startRange; i <= endRange; i++)
                {
                    result.Add(i.ToString());
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
        private static Level GetAboveLevel(Level[] levelColl, Level checkLevel, string levelSection)
        {
            string chkLvlNumber = GetLevelNumber(checkLevel);
            if (int.TryParse(chkLvlNumber, out int chkNumber))
            {
                if (chkNumber == 0)
                    throw new CheckerException($"У уровеня id {checkLevel.Id} значение уровня - 0. Это запрещено");

                foreach (Level level in levelColl)
                {
                    string lvlNumber = GetLevelNumber(level);
                    if (chkLvlNumber.Equals(lvlNumber)
                        || !GetLevelSections(level).Contains(levelSection))
                        continue;

                    if (int.TryParse(lvlNumber, out int number))
                    {
                        if (number == 0)
                            throw new CheckerException($"У уровеня id {level.Id} значение уровня - 0. Это запрещено");

                        if (chkNumber < 0
                            && number > 0
                            && chkNumber - number == -2
                            )
                            return level;
                        else if (number - chkNumber == 1)
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

        /// <summary>
        /// Взять уровень на 1 ниже по индексу этажа и номеру секции
        /// </summary>
        /// <param name="levelColl">Коллекция уровней для проверки</param>
        /// <param name="checkLevel">Уровень, по которому проверяем</param>
        /// <param name="levelSection">Номер секции</param>
        private static Level GetDownLevel(Level[] levelColl, Level checkLevel, string levelSection)
        {
            string chkLvlNumber = GetLevelNumber(checkLevel);
            if (int.TryParse(chkLvlNumber, out int chkNumber))
            {
                if (chkNumber == 0)
                    throw new CheckerException($"У уровеня id {checkLevel.Id} значение уровня - 0. Это запрещено");

                foreach (Level level in levelColl)
                {
                    string lvlNumber = GetLevelNumber(level);
                    if (chkLvlNumber.Equals(lvlNumber)
                        || !GetLevelSections(level).Contains(levelSection))
                        continue;

                    if (int.TryParse(lvlNumber, out int number))
                    {
                        if (number == 0)
                            throw new CheckerException($"У уровеня id {level.Id} значение уровня - 0. Это запрещено");

                        if (chkNumber > 0
                            && number < 0
                            && chkNumber - number == 2
                            )
                            return level;
                        else if (chkNumber - number == 1)
                            return level;
                    }
                    else
                        throw new CheckerException($"У уровеня id {level.Id} не удалось получить int его номера");
                }
            }
            else
                throw new CheckerException($"У уровеня id {checkLevel.Id} не удалось получить int его номера");

            return null;
        }
    }
}
