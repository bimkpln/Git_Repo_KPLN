using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;
using KPLN_Quantificator.Forms;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using Application = Autodesk.Navisworks.Api.Application;

namespace KPLN_Quantificator.Services
{
    internal static class ClashesCounter
    {
        private static readonly Document _doc = Application.ActiveDocument;


        private static List<string> _testNamesList = new List<string>
        {
            "АР",
            "КР",
            "ОВ",
            "ИТП",
            "ВК",
            "ПТ",
            "ЭОМ",
            "СС",
            "АВ",
            "ССАВ",
        };

        private static List<string> _detailedNamesList = new List<string>
        {
            "ОВ",
            "ИТП",
            "ВК",
            "ПТ",
            "ЭОМ",
            "СС",
            "АВ",
        };

        private static List<string> _ignoreNamesList = new List<string>
        {
            "АР",
            "КР"
        };

        private static Dictionary<string, int> _counterResult = new Dictionary<string, int>();
        private static Dictionary<string, int> _counterGroups = new Dictionary<string, int>();
        private static Dictionary<string, int> _counterNonGroups = new Dictionary<string, int>();

        private static Dictionary<string, int> _counterIosOnIos = new Dictionary<string, int>();
        private static Dictionary<string, int> _counterIosOnIosGroups = new Dictionary<string, int>();
        private static Dictionary<string, int> _counterIosOnIosNonGroups = new Dictionary<string, int>();

        private static List<string> _errorTestNames = new List<string>();

        public static void Prepare()
        {
            _counterResult.Clear();
            _counterGroups.Clear();
            _counterNonGroups.Clear();

            _counterIosOnIos.Clear();
            _counterIosOnIosGroups.Clear();
            _counterIosOnIosNonGroups.Clear();

            _errorTestNames.Clear();
        }













        /// <summary>
        /// Формирование данных по отчетам о коллизиях
        /// </summary>
        public static void Execute()
        {
            DocumentClashTests docClashTests = _doc.GetClash().TestsData;

            // Один Clash из ClashTest (из Item)
            foreach (SavedItem savedItem in docClashTests.Tests)
            {
                bool isExsist = false;

                // Дисплейное имя
                string testName = savedItem.DisplayName ?? string.Empty;

                // Список только ИОСных имён
                HashSet<string> detailedCodesInName = GetDetailedCodesInName(testName);

                // ИОС и (Дублирование)/(Самопересечение)
                bool nameHasDupOrSelf =
                    (testName.Contains("(Дублирование)") || testName.Contains("(Самопересечение)"))
                    && HasAnyDetailed(testName);

                // Отдел словоря со всеми отделами
                foreach (string namePart in _testNamesList)
                {

                    if (!IsCategoryMatch(testName, namePart))
                        continue;
                    isExsist = true;

                    int counterAll = 0;

                    int groupCounter = 0;
                    int nonGroupCounter = 0;

                    // Один Clash, как ClashTest
                    ClashTest currentClashTest = savedItem as ClashTest;
                    if (currentClashTest == null)
                    {
                        Output.PrintAlert($"Проблемы при определении отчета {testName}");
                        continue;
                    }

                    // Clash и группы Clash внутри отчёта
                    SavedItemCollection savedItemsColl = currentClashTest.Children;

                    foreach (SavedItem child in savedItemsColl)
                    {
                        if (child is ClashResult cr)
                        {
                            if (cr.Status == ClashResultStatus.Active || cr.Status == ClashResultStatus.New)
                            {
                                if (child.IsGroup)
                                    groupCounter++; // Групп
                                else
                                    nonGroupCounter++; // Без групп
                                counterAll++;
                            }
                        }

                        else if (child.IsGroup)
                        {
                            if (child is GroupItem gi)
                            {
                                bool hasRelevantClash = false;
                                foreach (SavedItem giChild in gi.Children)
                                {
                                    if (giChild is ClashResult crg)
                                    {
                                        if (crg.Status == ClashResultStatus.Active || crg.Status == ClashResultStatus.New)
                                        {
                                            hasRelevantClash = true;
                                            counterAll++;
                                        }
                                    }
                                    else
                                    {
                                        Output.PrintAlert($"Проблемы при определении отчета {giChild.DisplayName}");
                                    }
                                }
                                if (hasRelevantClash)
                                    groupCounter++; // Групп
                            }
                        }
                    }

                    // Всего
                    if (_counterResult.ContainsKey(namePart)) _counterResult[namePart] += counterAll;
                    else _counterResult.Add(namePart, counterAll);

                    // Групп
                    if (_counterGroups.ContainsKey(namePart)) _counterGroups[namePart] += groupCounter;
                    else _counterGroups.Add(namePart, groupCounter);

                    // Вне групп
                    if (_counterNonGroups.ContainsKey(namePart)) _counterNonGroups[namePart] += nonGroupCounter;
                    else _counterNonGroups.Add(namePart, nonGroupCounter);

                    if (_detailedNamesList.Contains(namePart))
                    {
                        // 1) игнор-отделы (АР/КР и т.п.)
                        bool hasIgnore = false;
                        foreach (var ign in _ignoreNamesList)
                        {
                            if (IsCategoryMatch(testName, ign)) { hasIgnore = true; break; }
                        }

                        // 2) фильтруем детальные коды через IsCategoryMatch
                        var filteredDetailed = new HashSet<string>();
                        foreach (var d in detailedCodesInName)
                        {
                            if (!IsCategoryMatch(testName, d)) continue;
                            filteredDetailed.Add(d);
                        }

                        // 3) составной "ССАВ/СС_АВ" детектим отдельно (он не в _detailedNamesList)
                        bool hasCompositeSSAV = IsCategoryMatch(testName, "ССАВ");

                        // 4) составной даёт «пару», только если нет игнора и есть хотя бы один детальный код
                        bool compositeGivesPair = hasCompositeSSAV && !hasIgnore && filteredDetailed.Count > 0;

                        // 5) берём в расчёт только если текущий детальный реально найден и нет игнора
                        bool eligible = filteredDetailed.Contains(namePart) && !hasIgnore;

                        if (eligible)
                        {
                            // 6) спец-кейс: “ОВ1 vs ОВ2” учитываем ТОЛЬКО в ИОС-на-ИОС
                            bool ovHasTwo = (namePart == "ОВ") && HasOV1andOV2(testName);

                            bool hasAnotherDetailed =
                                (filteredDetailed.Count > 1) ||                                  // другой детальный код
                                (filteredDetailed.Count == 1 && !filteredDetailed.Contains(namePart)) || // единственный код не равен namePart
                                compositeGivesPair ||                                             // пара от ССАВ/СС_АВ
                                ovHasTwo;                                                         // ОВ1 + ОВ2

                            bool isIosOnIos = hasAnotherDetailed || nameHasDupOrSelf;

                            if (isIosOnIos)
                            {
                                if (_counterIosOnIos.ContainsKey(namePart)) _counterIosOnIos[namePart] += counterAll;
                                else _counterIosOnIos.Add(namePart, counterAll);

                                if (_counterIosOnIosGroups.ContainsKey(namePart)) _counterIosOnIosGroups[namePart] += groupCounter;
                                else _counterIosOnIosGroups.Add(namePart, groupCounter);

                                if (_counterIosOnIosNonGroups.ContainsKey(namePart)) _counterIosOnIosNonGroups[namePart] += nonGroupCounter;
                                else _counterIosOnIosNonGroups.Add(namePart, nonGroupCounter);
                            }
                        }
                    }

                }

                if (!isExsist)
                    _errorTestNames.Add(savedItem.DisplayName);
            }



















            bool NameHasSSAV(string s) => !string.IsNullOrEmpty(s) && (s.Contains("ССАВ") || s.Contains("СС_АВ"));
            bool NameHasIgnore(string s)
            {
                if (string.IsNullOrEmpty(s)) return false;
                foreach (var ign in _ignoreNamesList)
                    if (!string.IsNullOrEmpty(ign) && s.Contains(ign))
                        return true;
                return false;
            }

            bool TestHasDupOrSelf(string s) => !string.IsNullOrEmpty(s) && (s.Contains("(Дублирование)") || s.Contains("(Самопересечение)"));

            var rxStandaloneDetailedInTestName = new System.Text.RegularExpressions.Regex(
                @"(?<![A-Za-zА-Яа-я0-9])" +
                @"(" +
                    @"ОВ(?:1|2)?|ИТП|ВК|ПТ|ЭОМ" +
                    @"|СС(?!_?АВ)" +
                    @"|(?<!СС)(?<!СС_)АВ" +
                @")" +
                @"(?![A-Za-zА-Яа-я0-9])",
                System.Text.RegularExpressions.RegexOptions.Compiled
            );
            bool TestHasAnyStandaloneDetailed(string s) => !string.IsNullOrEmpty(s) && rxStandaloneDetailedInTestName.IsMatch(s);

            bool IsCountableStatus(ClashResultStatus st) => st == ClashResultStatus.Active || st == ClashResultStatus.New;

            void Inc(Dictionary<string, int> dict, string key, int delta)
            {
                if (delta <= 0) return;
                if (dict.TryGetValue(key, out int cur)) dict[key] = cur + delta;
                else dict[key] = delta;
            }

            // Доп. проход: считаем ИОС-на-ИОС для составного "ССАВ/СС_АВ" в ключ "ССАВ"
            foreach (SavedItem savedItem in docClashTests.Tests)
            {
                string testName = savedItem.DisplayName ?? string.Empty;

                if (!NameHasSSAV(testName)) continue;
                if (NameHasIgnore(testName)) continue;

                bool compositeGivesPair = TestHasAnyStandaloneDetailed(testName); // Название нужное
                bool dupOrSelfInTest = TestHasDupOrSelf(testName); // Самопересечение
                bool isIosOnIosBySSAV = compositeGivesPair || dupOrSelfInTest;

                if (!isIosOnIosBySSAV)
                    continue;

                var currentClashTest = savedItem as ClashTest;
                if (currentClashTest == null) continue;

                int counterAll = 0, groupCounter = 0, nonGroupCounter = 0;

                var savedItemsColl = currentClashTest.Children;
                if (savedItemsColl != null && savedItemsColl.Count > 0)
                {
                    foreach (SavedItem child in savedItemsColl)
                    {
                        if (child is ClashResult cr)
                        {
                            if (IsCountableStatus(cr.Status))
                            {
                                if (child.IsGroup) groupCounter++;
                                else nonGroupCounter++;
                                counterAll++;
                            }
                        }
                        else if (child.IsGroup && child is GroupItem gi)
                        {
                            bool hasRelevant = false;
                            foreach (SavedItem giChild in gi.Children)
                            {
                                if (giChild is ClashResult crg && IsCountableStatus(crg.Status))
                                {
                                    hasRelevant = true;
                                    counterAll++;
                                }
                            }
                            if (hasRelevant) groupCounter++;
                        }
                    }
                }

                if (counterAll > 0)
                {
                    Inc(_counterIosOnIos, "ССАВ", counterAll);
                    Inc(_counterIosOnIosGroups, "ССАВ", groupCounter);
                    Inc(_counterIosOnIosNonGroups, "ССАВ", nonGroupCounter);
                }
            }
        }

        private static bool HasOV1andOV2(string s) => s.Contains("ОВ1") && s.Contains("ОВ2");

        // Единая проверка вхождения кода в названии теста
        private static bool IsCategoryMatch(string testName, string code)
        {
            switch (code)
            {
                case "СС":
                    // "СС", за которым НЕ идёт "АВ" и НЕ "_АВ"
                    return Regex.IsMatch(testName, @"СС(?!_?АВ)");

                case "АВ":
                    // "АВ", перед которым НЕ "СС" и НЕ "СС_"
                    return Regex.IsMatch(testName, @"(?<!СС)(?<!СС_)АВ");

                case "ССАВ":
                    // Ловим и слитно, и через подчёркивание
                    return Regex.IsMatch(testName, @"СС_?АВ");

                default:
                    // Остальные — по-старому
                    return testName.Contains(code);
            }
        }

        // Сбор детальных кодов
        private static HashSet<string> GetDetailedCodesInName(string testName)
        {
            var set = new HashSet<string>();
            foreach (var d in _detailedNamesList)
                if (IsCategoryMatch(testName, d))
                    set.Add(d);
            return set;
        }

        // Наличие ИОС-кода в имени
        private static bool HasAnyDetailed(string testName)
        {
            foreach (var d in _detailedNamesList)
                if (IsCategoryMatch(testName, d))
                    return true;
            return false;
        }

        /// <summary>
        /// Вывод результатов обработки отчетов о коллизиях
        /// </summary>
        public static void PrintResult()
        {
            if (_counterResult.Count > 0)
            {
                foreach (string key in _testNamesList)
                {
                    if (_counterResult.ContainsKey(key))
                    {
                        int clashCount = _counterResult[key];
                        int groupCount = _counterGroups.ContainsKey(key) ? _counterGroups[key] : 0;
                        int nonGroupCount = _counterNonGroups.ContainsKey(key) ? _counterNonGroups[key] : 0;

                        if (_detailedNamesList.Contains(key) || key == "ССАВ")
                        {
                            int iosTotal = _counterIosOnIos.ContainsKey(key) ? _counterIosOnIos[key] : 0;
                            int iosGroups = _counterIosOnIosGroups.ContainsKey(key) ? _counterIosOnIosGroups[key] : 0;
                            int iosNonGroups = _counterIosOnIosNonGroups.ContainsKey(key) ? _counterIosOnIosNonGroups[key] : 0;

                            Output.PrintSuccess(
                                $"Количество коллизий для раздела {key} составляет: {clashCount} шт. (Групп {groupCount + nonGroupCount}). " +
                                $"ИОС на ИОС - {iosTotal} шт. (Групп {iosGroups + iosNonGroups})",
                                boldKey: key,
                                boldTotalGroups: groupCount + nonGroupCount,
                                boldIosGroups: iosGroups + iosNonGroups
                            );
                        }
                        else
                        {
                            Output.PrintSuccess(
                                $"Количество коллизий для раздела {key} составляет: {clashCount} шт. (Групп {groupCount + nonGroupCount})",
                                boldKey: key,
                                boldTotalGroups: groupCount + nonGroupCount,
                                boldIosGroups: null
                            );
                        }
                    }
                }
            }

            if (_errorTestNames.Count > 0)
            {
                StringBuilder stringBuilder = new StringBuilder();
                foreach (string name in _testNamesList)
                {
                    stringBuilder.Append($" {name},");
                }
                Output.PrintAlert($"Имена отчетов должны содержать следующие имена разделов: {stringBuilder.ToString().TrimEnd(',')}");

                foreach (string name in _errorTestNames)
                {
                    Output.PrintAlert($"Ошибка в имени отчета: {name}");
                }
            }

            if (_counterResult.Count == 0 && _errorTestNames.Count == 0)
                Output.PrintAlert($"Не удалось обработать файл. Обратись к разработчику!");
        }
    }
}
