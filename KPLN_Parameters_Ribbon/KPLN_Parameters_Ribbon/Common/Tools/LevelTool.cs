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
        
        private static IEnumerable<Level> _levels = null;

        /// <summary>
        /// Коллекция уровней в проекте
        /// </summary>
        public static IEnumerable<Level> Levels
        {
            get { return _levels; }
            set { _levels = value; }
        }

        /// <summary>
        /// Словарь "номер уровня - уровень". Для экономии ресурсов
        /// </summary>
        private static Dictionary<ElementId, string> _levelNumberMap = new Dictionary<ElementId, string>();

        /// <summary>
        /// Поиск номера уровня у элемента (над уровнем) по уровню
        /// </summary>
        /// <param name="lev">Уровень</param>
        /// <param name="floorTextPosition">Индекс номера уровня с учетом разделителя</param>
        /// <param name="splitChar">Разделитель</param>
        /// <returns>Номер уровня</returns>
        public static string GetFloorNumberByLevel(Level lev, int floorTextPosition, char splitChar)
        {
            if (_levelNumberMap.ContainsKey(lev.Id))
            {
                return _levelNumberMap[lev.Id];
            }
            else
            {
                string levname = lev.Name;
                if (levname.ToLower().Contains("кровля"))
                {
                    _levelNumberMap.Add(lev.Id, "99");
                    return "99";
                }

                string[] splitname = levname.Split(splitChar);
                if (splitname.Length < 2)
                    throw new Exception($"Некорректное имя уровня: {levname}");

                string floorNumber = splitname[floorTextPosition];

                // Это исключительно для Обыденского
                if (floorNumber.Contains("0") && !floorNumber.EndsWith("0"))
                    floorNumber = floorNumber.Replace("0", "");

                if (floorNumber.Contains("-"))
                    floorNumber = floorNumber.Replace("-", "м");

                _levelNumberMap.Add(lev.Id, floorNumber);
                return floorNumber;
            }
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
            if (levname.ToLower().Contains("кровля"))
            {
                _levelNumberMap.Add(lev.Id, "99");
                return "99";
            }

            string[] splitname = levname.Split(splitChar);
            if (splitname.Length < 2)
                throw new Exception($"Некорректное имя уровня: {levname}");

            string floorNumber = splitname[floorTextPosition];

            // Это исключительно для Обыденского
            if (floorNumber.Contains("0") && !floorNumber.EndsWith("0"))
                floorNumber = floorNumber.Replace("0", "");

            if (floorNumber.Contains("-"))
                floorNumber = floorNumber.Replace("-", "м");

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
        public static string GetFloorNumberDecrementLevel(Level lev, int floorTextPosition, char splitChar)
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
                levId = GetNearestBelowPointLevel(minPoint, doc).Id;
            }

            if (levId == ElementId.InvalidElementId)
                throw new Exception("Не удалось получить уровень у элемента с Id:" + elem.Id.IntegerValue.ToString());

            Level lev = doc.GetElement(levId) as Level;
            return lev;
        }

        /// <summary>
        /// Поиск уровня, ближайшего к точке и ниже этой точки. Для элементов на нижних уровнях с отрицательной отметкой выдаст нижний уровень
        /// </summary>
        /// <param name="point">Точка в пространстве</param>
        /// <param name="doc">Документ для анализа</param>
        /// <returns>Id уровня</returns>
        public static Level GetNearestBelowPointLevel(XYZ point, Document doc)
        {
            double pointZ = point.Z;

            BasePoint projectBasePoint = new FilteredElementCollector(doc)
                .OfClass(typeof(BasePoint))
                .WhereElementIsNotElementType()
                .Cast<BasePoint>()
                .Where(i => i.IsShared == false)
                .First();
            double projectPointElevation = projectBasePoint.get_BoundingBox(null).Min.Z;
            
            Level nearestLevel = null;

            double offset = 10000;
            foreach (Level lev in Levels)
            {
                double levHeigth = lev.Elevation + projectPointElevation;
                double tempElev = pointZ - levHeigth;
                if (tempElev < 0) continue;

                if (tempElev < offset)
                {
                    nearestLevel = lev;
                    offset = tempElev;
                }
            }

            if (nearestLevel == null)
            {
                foreach (Level lev in Levels)
                {
                    if (nearestLevel == null)
                    {
                        nearestLevel = lev;
                        continue;
                    }
                    if (lev.Elevation < nearestLevel.Elevation)
                    {
                        nearestLevel = lev;
                    }
                }
            }

            return nearestLevel;
        }

        /// <summary>
        /// Поиск уровня, ближайшего к уровню и ниже его. Для элементов на нижних уровнях с отрицательной отметкой выдаст нижний уровень
        /// </summary>
        /// <param name="point">Точка в пространстве</param>
        /// <param name="doc">Документ для анализа</param>
        /// <returns>Id уровня</returns>
        public static Level GetNearestBelowLevel(Level level, Document doc)
        {
            double zValue = level.Elevation;
            return GetNearestBelowPointLevel(new XYZ(0, 0, zValue), doc);
        }

        /// <summary>
        /// Поиск уровня, ближайшего к уровню и выше его
        /// </summary>
        /// <param name="point">Точка в пространстве</param>
        /// <param name="doc">Документ для анализа</param>
        /// <returns>Id уровня</returns>
        private static Level GetNearestUpperLevel(Level level, Document doc)
        {
            double zValue = level.Elevation;
            
            return Levels
                .Where(x => x.Elevation > zValue || x.Elevation == zValue)
                .OrderBy(x => x.Elevation)
                .FirstOrDefault();
        }

        /// <summary>
        /// Поиск значения привязки к уровню у элемента
        /// </summary>
        /// <param name="elem">Элемент, у которого осуществляется поиск</param>
        /// <returns>Id уровня</returns>
        public static double GetElementLevelGrip(Element elem, Level level)
        {
            // Первичный отлов по категории
            double ZCoord;
            Type elemType = elem.GetType();
            switch (elemType.Name)
            {
                case nameof(ExtrusionRoof):
                    double offsetRoofConstr = elem.get_Parameter(BuiltInParameter.ROOF_CONSTRAINT_OFFSET_PARAM).AsDouble();
                    return offsetRoofConstr;
                case nameof(FootPrintRoof):
                    double offsetRoof = elem.get_Parameter(BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM).AsDouble();
                    return offsetRoof;
                case nameof(Railing):
                    double offsetStairs = elem.get_Parameter(BuiltInParameter.STAIRS_RAILING_HEIGHT_OFFSET).AsDouble();
                    return offsetStairs;
            }

            // Вторичный отлов по типу Loaction
            Location location = elem.Location;
            Type locationType = location.GetType();
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
                Level upperLev = null;
                foreach (Level lvl in Levels)
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
