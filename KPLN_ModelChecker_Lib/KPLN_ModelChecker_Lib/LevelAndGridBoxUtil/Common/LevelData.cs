using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace KPLN_ModelChecker_Lib.LevelAndGridBoxUtil.Common
{
    public class LevelData
    {
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
                        minPointOfLevels = CurrentLevel.Elevation - 3;
                    else
                        minPointOfLevels = CurrentLevel.Elevation;

                    double maxPointOLevels;
                    if (CurrentAboveLevel == null)
                        maxPointOLevels = minPointOfLevels + 3;
                    else
                        maxPointOLevels = CurrentAboveLevel.Elevation;

                    _minAndMaxLvlPnts[0] = minPointOfLevels;
                    _minAndMaxLvlPnts[1] = maxPointOLevels;
                }

                return _minAndMaxLvlPnts;
            }
        }

        /// <summary>
        /// Номер уровня
        /// </summary>
        public string CurrentLevelNumber { get; private set; }

        /// <summary>
        /// Секции для уровня
        /// </summary>
        public string CurrentSectionNumber { get; private set; }

        private LevelData(Level level, string lvlNumber, string sectNumber, Level aboveLevel, Level downLevel)
        {
            CurrentLevel = level;
            CurrentLevelNumber = lvlNumber;
            CurrentSectionNumber = sectNumber;
            CurrentAboveLevel = aboveLevel;
            CurrentDownLevel = downLevel;
        }

        /// <summary>
        /// Подготовка коллекции уровней для анализа
        /// </summary>
        /// <param name="doc">Revit-документ для анализа</param>
        internal static List<LevelData> LevelPrepare(Document doc)
        {
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
                    LevelData myLevel = new LevelData(level, lvlNumber, levelSection, aboveLevel, downLevel);
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

            if (!splitname[0].Contains("С") && !splitname[0].Contains("К") && !splitname[0].Contains("ПАР") && !splitname[0].Contains("СТЛ"))
                throw new CheckerException($"Некорректное имя уровня (см. ВЕР, кодировка секции/корпуса: С1/С1.2/С1-6/К1/ПАР/СТЛ. Id: {level.Id}");

            return GetNumbers(splitname[0]);
        }

        private static List<string> GetNumbers(string input)
        {
            List<string> numbers = new List<string>();
            MatchCollection matches = Regex.Matches(input, @"\d+");
            foreach (Match match in matches)
            {
                numbers.Add(match.Value);
            }
            return numbers;
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
