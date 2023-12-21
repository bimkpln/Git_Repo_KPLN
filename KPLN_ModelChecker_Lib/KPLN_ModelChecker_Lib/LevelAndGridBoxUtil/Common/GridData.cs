using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_Lib.LevelAndGridBoxUtil.Common
{
    public class GridData
    {
        /// <summary>
        /// Коллеция осей для секции
        /// </summary>
        public HashSet<Grid> CurrentGrids { get; private set; }

        /// <summary>
        /// Имя секции
        /// </summary>
        public string CurrentSection { get; private set; }

        private GridData(string currentSection, HashSet<Grid> currentGrids)
        {
            CurrentSection = currentSection;
            CurrentGrids = currentGrids;
        }

        /// <summary>
        /// Подготовка коллекции осей для анализа
        /// </summary>
        /// <param name="doc">Revit-документ для анализа</param>
        /// <param name="paramName">Имя параметра для сепарации</param>
        internal static List<GridData> GridPrepare(Document doc, string paramName)
        {
            List<GridData> preapareGrids = new List<GridData>();

            Grid[] grids = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(Grid))
                .Cast<Grid>()
                .ToArray();
            foreach (Grid grid in grids)
            {
                Parameter param = grid.LookupParameter(paramName);
                if (param != null && param.AsString() != null && param.AsString().Length != 0)
                {
                    foreach (string sect in param.AsString().Split('-'))
                    {
                        List<GridData> equalSections = preapareGrids.Where(g => g.CurrentSection.Equals(sect)).ToList();
                        if (equalSections.Count > 0)
                        {
                            foreach (GridData gd in equalSections)
                            {
                                gd.CurrentGrids.Add(grid);
                            }
                        }
                        else
                        {
                            preapareGrids.Add(new GridData(sect, new HashSet<Grid>() { grid }));
                        }
                    }
                }
            }

            // Проверка полученных данных
            foreach (GridData gd in preapareGrids)
            {
                if (gd.CurrentGrids.Count < 4)
                    throw new CheckerException($"Количество осей с номером секции: {gd.CurrentSection} меньше 4. Проверьте назначение параметров у осей!");
            }

            if (preapareGrids.Count == 0)
                throw new CheckerException($"Для заполнения номера секции в элементах, необходимо заполнить параметр: {paramName} в осях! Значение указывается через \"-\" для осей, относящихся к нескольким секциям.");

            return preapareGrids;
        }
    }
}