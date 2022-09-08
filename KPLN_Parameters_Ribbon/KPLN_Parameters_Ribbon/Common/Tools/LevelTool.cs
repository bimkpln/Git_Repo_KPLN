using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Loader.Output.Output;

namespace KPLN_Parameters_Ribbon.Common.Tools
{
    /// <summary>
    /// Инструмент анализа уровней
    /// </summary>
    internal static class LevelTool
    {
        /// <summary>
        /// Максимальный уровень. Подходит только для 1 секционных зданий!!!! Переделать на поиск мах уровней для каждой секции
        /// </summary>
        private static Level _maxLevel = null;

        /// <summary>
        /// Верхний уровень, у которого допускается отсутсвие парамтера "На уровень выше"
        /// </summary>
        public static Level MaxLevel
        {
            get
            {
                return _maxLevel;
            }
            private set
            {
                _maxLevel = value;
            }
        }

        /// <summary>
        /// Поиск номера уровня у элемента (над уровнем) по уровню
        /// </summary>
        /// <param name="lev">Уровень</param>
        /// <param name="floorTextPosition">Индекс номера уровня с учетом разделителя</param>
        /// <param name="splitChar">Разделитель</param>
        /// <returns>Номер уровня</returns>
        public static string GetFloorNumberByLevel(Level lev, int floorTextPosition, char splitChar)
        {
            string levname = lev.Name;
            string[] splitname = levname.Split(splitChar);
            if (splitname.Length < 2)
            {
                Print("Некорректное имя уровня: " + levname, KPLN_Loader.Preferences.MessageType.Error);
                return null;
            }
            string floorNumber = splitname[floorTextPosition];
            return floorNumber;
        }

        /// <summary>
        /// Поиск номера уровня у элемента (под уровнем) по уровню выше
        /// </summary>
        /// <param name="lev">Уровень</param>
        /// <param name="floorTextPosition">Индекс номера уровня с учетом разделителя</param>
        /// <param name="doc">Документ для анализа</param>
        /// <param name="splitChar">Разделитель</param>
        /// <returns>Номер уровня</returns>
        public static string GetFloorNumberIncrementLevel(Level lev, int floorTextPosition, Document doc, char splitChar)
        {
            SetMaxLevel(doc);

            ElementId aboveLevelId = lev.get_Parameter(BuiltInParameter.LEVEL_UP_TO_LEVEL).AsElementId();
            Level aboveLevel = (Level)doc.GetElement(aboveLevelId);
            if (aboveLevel == null && !lev.Id.Equals(MaxLevel.Id))
            {
                return null;
            }
            else if (lev.Id.Equals(MaxLevel.Id))
            {
                aboveLevel = MaxLevel;
            }

            string levname = aboveLevel.Name;
            string[] splitname = levname.Split(splitChar);
            if (splitname.Length < 2)
            {
                Print("Некорректное имя уровня: " + levname, KPLN_Loader.Preferences.MessageType.Error);
                return null;
            }
            string floorNumber = splitname[floorTextPosition];
            return floorNumber;
        }

        /// <summary>
        /// Поиск номера уровня у элемента (под уровнем) по уровню ниже
        /// </summary>
        /// <param name="lev">Уровень</param>
        /// <param name="floorTextPosition">Индекс номера уровня с учетом разделителя</param>
        /// <param name="doc">Документ для анализа</param>
        /// <param name="splitChar">Разделитель</param>
        /// <returns>Номер уровня</returns>
        public static string GetFloorNumberDecrementLevel(Level lev, int floorTextPosition, Document doc, char splitChar)
        {
            List<char> resultFloorNumber = new List<char>();
            string floorNumber = GetFloorNumberByLevel(lev, floorTextPosition, splitChar);
            foreach(char c in floorNumber)
            {
                if (c == '0' || c == '-')
                {
                    resultFloorNumber.Add(c);
                }
                else if (floorNumber.Contains('-'))
                {
                    if (Int32.TryParse(c.ToString(), out int intC))
                    {
                        resultFloorNumber.Add((++intC).ToString()[0]);
                    }
                }
                else
                {
                    if (Int32.TryParse(c.ToString(), out int intC))
                    {
                        resultFloorNumber.Add((--intC).ToString()[0]);
                    }
                }
            }
            return new string(resultFloorNumber.ToArray());
        }

        /// <summary>
        /// Поиск уровня у элемента
        /// </summary>
        /// <param name="elem">Элемент</param>
        /// <param name="doc">Документ для анализа</param>
        /// <returns>Уровень</returns>
        /// <exception cref="Exception"></exception>
        public static Level GetLevelOfElement(Element elem, Document doc)
        {
            ElementId levId = elem.LevelId;
            if (levId == ElementId.InvalidElementId)
            {
                Parameter levParam = elem.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM);
                if (levParam != null)
                    levId = levParam.AsElementId();
            }

            if (levId == ElementId.InvalidElementId)
            {
                Parameter levParam = elem.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM);
                if (levParam != null)
                    levId = levParam.AsElementId();
            }

            if (levId == ElementId.InvalidElementId)
            {
                Parameter levParam = elem.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM);
                if (levParam != null)
                    levId = levParam.AsElementId();
            }

            if (levId == ElementId.InvalidElementId && elem is StairsRun)
            {
                StairsRun run = elem as StairsRun;
                if (run != null)
                {
                    Stairs stair = run.GetStairs();
                    Parameter levParam = stair.get_Parameter(BuiltInParameter.STAIRS_BASE_LEVEL_PARAM);
                    if (levParam != null)
                        levId = levParam.AsElementId();
                }
            }

            if (levId == ElementId.InvalidElementId)
            {
                List<Solid> solids = GeometryTool.GetSolidsFromElement(elem);
                if (solids.Count == 0) return null;
                XYZ[] maxmin = GeometryTool.GetMaxMinHeightPoints(solids);
                XYZ minPoint = maxmin[1];
                levId = GetNearestLevel(minPoint, doc);
            }

            if (levId == ElementId.InvalidElementId)
                throw new Exception("Не удалось получить уровень у элемента с Id:" + elem.Id.IntegerValue.ToString());

            Level lev = doc.GetElement(levId) as Level;
            return lev;
        }

        /// <summary>
        /// Поиск ближайшего уровня
        /// </summary>
        /// <param name="point">Точка в пространстве</param>
        /// <param name="doc">Документ для анализа</param>
        /// <returns>Id уровня</returns>
        public static ElementId GetNearestLevel(XYZ point, Document doc)
        {
            BasePoint projectBasePoint = new FilteredElementCollector(doc)
                .OfClass(typeof(BasePoint))
                .WhereElementIsNotElementType()
                .Cast<BasePoint>()
                .Where(i => i.IsShared == false)
                .First();
            double projectPointElevation = projectBasePoint.get_BoundingBox(null).Min.Z;

            double pointZ = point.Z;
            List<Level> levels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .WhereElementIsNotElementType()
                .Cast<Level>()
                .ToList();

            Level finalLevel = null;

            foreach (Level lev in levels)
            {
                if (finalLevel == null)
                {
                    finalLevel = lev;
                    continue;
                }
                if (lev.Elevation < finalLevel.Elevation)
                {
                    finalLevel = lev;
                    continue;
                }
            }

            double offset = 10000;
            foreach (Level lev in levels)
            {
                double levHeigth = lev.Elevation + projectPointElevation;
                double testElev = pointZ - levHeigth;
                if (testElev < 0) continue;

                if (testElev < offset)
                {
                    finalLevel = lev;
                    offset = testElev;
                }
            }

            return finalLevel.Id;
        }

        /// <summary>
        /// Поиск значения привязки к уровню у элемента
        /// </summary>
        /// <param name="elem">Элемент, у которого осуществляется поиск</param>
        /// <returns>Id уровня</returns>
        public static double GetElementLevelGrip(Element elem, Level level)
        {
            Location location = elem.Location;
            Type locationType = location.GetType();

            double ZCoord;
            switch (locationType.Name)
            {
                case nameof(LocationCurve):
                    LocationCurve locCurve = location as LocationCurve;
                    Curve curve = locCurve.Curve;
                    ZCoord = Math.Abs((curve.GetEndPoint(0) - curve.GetEndPoint(1)).Z);
                    return ZCoord - level.Elevation;
                case nameof(Location):
                    double offsetHeightAbove = elem.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).AsDouble();
                    return offsetHeightAbove;
                case nameof(LocationPoint):
                    LocationPoint locationPoint = location as LocationPoint;
                    ZCoord = locationPoint.Point.Z;
                    return ZCoord - level.Elevation;
                default:
                    throw new Exception("Не удалось получить отметку у элемента с Id:" + elem.Id.IntegerValue.ToString());
            }
        }

        /// <summary>
        /// Поиск верхнего уровня
        /// </summary>
        /// <returns></returns>
        private static Level SetMaxLevel (Document doc)
        {
            if (MaxLevel == null)
            {
                var levColl = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Levels)
                    .WhereElementIsNotElementType();

                Level upperLev = null;
                foreach (Level lvl in levColl)
                {
                    if (upperLev == null)
                    {
                        upperLev = lvl;
                    }
                    else if (lvl.Elevation > upperLev.Elevation)
                    {
                        upperLev = lvl;
                    }
                }
                MaxLevel = upperLev;
            }

            return MaxLevel;
        }
    }
}
