using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;
using KPLN_Quantificator.Forms;
using System.Collections.Generic;
using System.Text;

namespace KPLN_Quantificator.Services
{
    internal static class ClashesCounter
    {
        private static readonly Document _doc = Application.ActiveDocument;

        private static List<string> _testNames = new List<string>
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

        private static List<string> _detailedNames = new List<string>
        {
            "ОВ",
            "ИТП",
            "ВК",
            "ПТ",
            "ЭОМ",
            "СС",
            "АВ",
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
            _errorTestNames.Clear();

            _counterIosOnIos.Clear();
            _counterIosOnIosGroups.Clear();
            _counterIosOnIosNonGroups.Clear();
        }

        /// <summary>
        /// Формирование данных по отчетам о коллизиях
        /// </summary>
        public static void Execute()
        {
            DocumentClashTests docClashTests = _doc.GetClash().TestsData;

            foreach (SavedItem savedItem in docClashTests.Tests)
            {
                bool isExsist = false;
                string testName = savedItem.DisplayName ?? string.Empty;
                bool isSSAV = savedItem.DisplayName.Contains("СС") && savedItem.DisplayName.Contains("АВ");

                HashSet<string> detailedCodesInName = new HashSet<string>();
                foreach (var d in _detailedNames)
                    if (testName.Contains(d))
                        detailedCodesInName.Add(d);

                bool nameHasDupOrSelf = testName.Contains("(Дублирование)") || testName.Contains("(Самопересечение)");

                foreach (string namePart in _testNames)
                {
                    if (!testName.Contains(namePart))
                        continue;

                    if (namePart == "СС" && isSSAV)
                        continue;

                    isExsist = true;

                    int counterAll = 0;
                    int groupCounter = 0;
                    int nonGroupCounter = 0;

                    ClashTest currentClashTest = savedItem as ClashTest;
                    if (currentClashTest == null)
                    {
                        Output.PrintAlert($"Проблемы при определении отчета {testName}");
                        continue;
                    }

                    SavedItemCollection savedItemsColl = currentClashTest.Children;
                    foreach (SavedItem child in savedItemsColl)
                    {
                        if (child is ClashResult cr)
                        {
                            if (cr.Status == ClashResultStatus.Active || cr.Status == ClashResultStatus.New)
                            {
                                if (child.IsGroup)
                                    groupCounter++;
                                else
                                    nonGroupCounter++;
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
                                    groupCounter++;
                            }
                        }
                    }

                    if (_counterResult.ContainsKey(namePart)) _counterResult[namePart] += counterAll;
                    else _counterResult.Add(namePart, counterAll);

                    if (_counterGroups.ContainsKey(namePart)) _counterGroups[namePart] += groupCounter;
                    else _counterGroups.Add(namePart, groupCounter);

                    if (_counterNonGroups.ContainsKey(namePart)) _counterNonGroups[namePart] += nonGroupCounter;
                    else _counterNonGroups.Add(namePart, nonGroupCounter);

                    if (_detailedNames.Contains(namePart))
                    {
                        bool hasAnotherDetailed =
                            detailedCodesInName.Count > 1 ||
                            (detailedCodesInName.Count == 1 && !detailedCodesInName.Contains(namePart)); 

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
                                                        
                if (!isExsist)
                    _errorTestNames.Add(savedItem.DisplayName);
            }
        }





        /// <summary>
        /// Вывод результатов обработки отчетов о коллизиях
        /// </summary>
        public static void PrintResult()
        {
            if (_counterResult.Count > 0)
            {
                foreach (string key in _testNames)
                {
                    if (_counterResult.ContainsKey(key))
                    {
                        int clashCount = _counterResult[key];
                        int groupCount = _counterGroups.ContainsKey(key) ? _counterGroups[key] : 0;
                        int nonGroupCount = _counterNonGroups.ContainsKey(key) ? _counterNonGroups[key] : 0;

                        if (_detailedNames.Contains(key))
                        {
                            int iosTotal = _counterIosOnIos.ContainsKey(key) ? _counterIosOnIos[key] : 0;
                            int iosGroups = _counterIosOnIosGroups.ContainsKey(key) ? _counterIosOnIosGroups[key] : 0;
                            int iosNonGroups = _counterIosOnIosNonGroups.ContainsKey(key) ? _counterIosOnIosNonGroups[key] : 0;

                            Output.PrintSuccess(
                                $"Количество коллизий для раздела {key} составляет: {clashCount} шт. (Групп {groupCount + nonGroupCount}). " +
                                $"ИОС на ИОС - {iosTotal} шт. (Групп {iosGroups + iosNonGroups})"
                            );
                        }
                        else
                        {
                            Output.PrintSuccess(
                                $"Количество коллизий для раздела {key} составляет: {clashCount} шт. (Групп {groupCount + nonGroupCount})"
                            );
                        }
                    }
                }
            }

            if (_errorTestNames.Count > 0)
            {
                StringBuilder stringBuilder = new StringBuilder();
                foreach (string name in _testNames)
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
