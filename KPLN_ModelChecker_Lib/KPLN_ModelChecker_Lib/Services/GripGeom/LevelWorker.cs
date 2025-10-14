using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;

namespace KPLN_ModelChecker_Lib.Services.GripGeom
{
    public static class LevelWorker
    {
        /// <summary>
        /// Бинарный поиск уровня по отметке 
        /// </summary>
        /// <param name="levels">Отсортированный список уровней</param>
        /// <param name="elevation">Отметка, по которой идёт поиск</param>
        /// <returns></returns>
        public static Level BinaryFindExactLevel(List<Level> levels, double elevation)
        {
            int low = 0;
            int high = levels.Count - 1;

            const double epsilon = 1e-6; // каб улічыць дробныя розніцы Revit (футаў)

            while (low <= high)
            {
                int mid = (low + high) / 2;


#if Debug2020 || Revit2020
                double midElev = UnitUtils.ConvertFromInternalUnits(levels[mid].Elevation, DisplayUnitType.DUT_MILLIMETERS);
#else
                double midElev = UnitUtils.ConvertFromInternalUnits(levels[mid].Elevation, SpecTypeId.Length);
#endif

                double diff = midElev - elevation;

                if (Math.Abs(diff) < epsilon)
                    return levels[mid]; // знайшлі дакладна

                if (diff < 0)
                    low = mid + 1;
                else
                    high = mid - 1;
            }

            return null; // дакладнага ўзроўню няма
        }
    }
}
