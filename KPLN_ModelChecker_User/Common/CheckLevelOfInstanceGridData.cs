using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_User.Common
{
    internal class CheckLevelOfInstanceGridData
    {
        private readonly static string _paramName = "КП_О_Секция";

        /// <summary>
        /// Коллеция осей для секции
        /// </summary>
        public HashSet<Grid> CurrentGrids { get; private set; }

        /// <summary>
        /// Имя секции
        /// </summary>
        public string CurrentSection { get; private set; }

        private CheckLevelOfInstanceGridData(string currentSection, HashSet<Grid> currentGrids)
        {
            CurrentSection = currentSection;
            CurrentGrids = currentGrids;
        }

        /// <summary>
        /// Подготовка коллекции осей для анализа
        /// </summary>
        /// <param name="doc">Revit-документ для анализа</param>
        public static List<CheckLevelOfInstanceGridData> GridPrepare(Document doc)
        {
            List<CheckLevelOfInstanceGridData> preapareGrids = new List<CheckLevelOfInstanceGridData>();

            Grid[] grids = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(Grid))
                .Cast<Grid>()
                .ToArray();
            foreach (Grid grid in grids)
            {
                Parameter param = grid.LookupParameter(_paramName);
                if (param != null && param.AsString() != null && param.AsString().Length != 0)
                {
                    foreach (string sect in param.AsString().Split('-'))
                    {
                        List<CheckLevelOfInstanceGridData> equalSections = preapareGrids.Where(g => g.CurrentSection.Equals(sect)).ToList();
                        if (equalSections.Count > 0)
                        {
                            foreach (CheckLevelOfInstanceGridData gd in equalSections)
                            {
                                gd.CurrentGrids.Add(grid);
                            }
                        }
                        else
                        {
                            preapareGrids.Add(new CheckLevelOfInstanceGridData(sect, new HashSet<Grid>() { grid }));
                        }
                    }
                }
            }

            // Проверка полученных данных
            foreach (CheckLevelOfInstanceGridData gd in preapareGrids)
            {
                if (gd.CurrentGrids.Count < 4)
                    throw new UserException($"Количество осей с номером секции: {gd.CurrentSection} меньше 4. Проверьте назначение параметров у осей!");
            }

            if (preapareGrids.Count == 0)
                throw new UserException($"Для заполнения номера секции в элементах, необходимо заполнить параметр: {_paramName} в осях! Значение указывается через \"-\" для осей, относящихся к нескольким секциям.");

            return preapareGrids;
        }
    }
}
