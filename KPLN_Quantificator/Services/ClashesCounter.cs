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

        private static Dictionary<string, int> _counterResult = new Dictionary<string, int>();
        private static Dictionary<string, int> _counterGroups = new Dictionary<string, int>();
        private static Dictionary<string, int> _counterNonGroups = new Dictionary<string, int>(); 

        private static List<string> _errorTestNames = new List<string>();

        public static void Prepare()
        {
            _counterResult.Clear();
            _counterGroups.Clear();
            _counterNonGroups.Clear();
            _errorTestNames.Clear();
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
                bool isSSAV = savedItem.DisplayName.Contains("СС") && savedItem.DisplayName.Contains("АВ");

                foreach (string namePart in _testNames)
                {
                    if (savedItem.DisplayName.Contains(namePart))
                    {
                        if (namePart == "СС" && isSSAV)
                            continue;

                        isExsist = true;

                        int counterAll = 0;
                        int groupCounter = 0;
                        int nonGroupCounter = 0;

                        ClashTest currentClashTest = savedItem as ClashTest;

                        if (currentClashTest != null)
                        {
                            SavedItemCollection savedItemsColl = currentClashTest.Children;

                            foreach (SavedItem savedItemInClashTest in savedItemsColl)
                            {
                                ClashResult clashResult = savedItemInClashTest as ClashResult;

                                if (clashResult != null)
                                {
                                    if (clashResult.Status == ClashResultStatus.Active || clashResult.Status == ClashResultStatus.New)
                                    {
                                        if (savedItemInClashTest.IsGroup)
                                        {
                                            groupCounter++;
                                        }
                                        else
                                        {
                                            nonGroupCounter++;
                                        }

                                        counterAll++;
                                    }
                                }
                                else if (savedItemInClashTest.IsGroup)
                                {
                                    GroupItem groupItem = savedItemInClashTest as GroupItem;

                                    if (groupItem != null)
                                    {
                                        bool hasRelevantClash = false;

                                        foreach (SavedItem savedItemInClashTestInGroup in groupItem.Children)
                                        {
                                            ClashResult clashResultInGroup = savedItemInClashTestInGroup as ClashResult;

                                            if (clashResultInGroup != null)
                                            {
                                                if (clashResultInGroup.Status == ClashResultStatus.Active || clashResultInGroup.Status == ClashResultStatus.New)
                                                {
                                                    hasRelevantClash = true;
                                                    counterAll++;
                                                }
                                            }
                                            else
                                            {
                                                Output.PrintAlert($"Проблемы при определении отчета {savedItemInClashTestInGroup.DisplayName}");
                                            }
                                        }

                                        if (hasRelevantClash)
                                        {
                                            groupCounter++;
                                        }
                                    }
                                }
                            }

                            if (_counterResult.ContainsKey(namePart))
                                _counterResult[namePart] += counterAll;
                            else
                                _counterResult.Add(namePart, counterAll);

                            if (_counterGroups.ContainsKey(namePart))
                                _counterGroups[namePart] += groupCounter;
                            else
                                _counterGroups.Add(namePart, groupCounter);

                            if (_counterNonGroups.ContainsKey(namePart))
                                _counterNonGroups[namePart] += nonGroupCounter;
                            else
                                _counterNonGroups.Add(namePart, nonGroupCounter);
                        }
                        else
                        {
                            Output.PrintAlert($"Проблемы при определении отчета {savedItem.DisplayName}");
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

                        Output.PrintSuccess($"Количество коллизий для раздела {key} составляет: {clashCount} шт. (Групп {groupCount + nonGroupCount})");
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
