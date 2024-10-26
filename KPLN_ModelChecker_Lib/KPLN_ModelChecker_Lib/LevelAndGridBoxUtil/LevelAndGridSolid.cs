using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using KPLN_ModelChecker_Lib.LevelAndGridBoxUtil.Common;

namespace KPLN_ModelChecker_Lib.LevelAndGridBoxUtil
{
    /// <summary>
    /// Класс для генерации солида между уровнем и ограждающими секциями
    /// </summary>
    public class LevelAndGridSolid
    {
        /// <summary>
        /// Солид в границах уровней и осей
        /// </summary>
        public Solid CurrentSolid { get; private set; }

        /// <summary>
        /// Ссылка на текущий CheckLevelOfInstanceLevelData
        /// </summary>
        public LevelData CurrentLevelData { get; private set; }

        /// <summary>
        /// Ссылка на CheckLevelOfInstanceGridData
        /// </summary>
        public GridData GridData { get; private set; }

        private LevelAndGridSolid(Solid solid, LevelData currentLevel, GridData gData)
        {
            CurrentSolid = solid;
            CurrentLevelData = currentLevel;
            GridData = gData;
        }

        /// <summary>
        /// Подготовка коллекции солидов с данными по секциям
        /// </summary>
        /// <param name="doc">Revit-документ</param>
        /// <param name="sectSeparParamName">Параметр для разделения осей и уровней по секциям</param>
        /// <param name="levelIndexParamName">Параметр для разделения уровней по этажам</param>
        /// <param name="floorScreedHeight">Толщина стяжки пола АР</param>
        /// <param name="downAndTopExtra">Расширение границ для самого нижнего и самого верхнего уровней</param>
        public static List<LevelAndGridSolid> PrepareSolids(Document doc, string sectSeparParamName,
            string levelIndexParamName, double floorScreedHeight = 0, double downAndTopExtra = 3)
        {
            List<LevelAndGridSolid> result = new List<LevelAndGridSolid>();

            List<GridData> gridDatas = GridData.GridPrepare(doc, sectSeparParamName);
            HashSet<string> multiGridsSet = new HashSet<string>(gridDatas.Select(g => g.CurrentSection));
            List<LevelData> levelDatas = multiGridsSet.Count == 1
                ? LevelData.LevelPrepare(doc, floorScreedHeight, downAndTopExtra, sectSeparParamName,
                    levelIndexParamName, multiGridsSet)
                : LevelData.LevelPrepare(doc, floorScreedHeight, downAndTopExtra, sectSeparParamName,
                    levelIndexParamName);

            // Подготовка предварительной коллекции элементов
            List<LevelAndGridSolid> preResult = new List<LevelAndGridSolid>();
            foreach (LevelData currentLevel in levelDatas)
            {
                foreach (GridData gData in gridDatas)
                {
                    if (currentLevel.CurrentSectionNumber.Equals(gData.CurrentSection))
                    {
                        Solid levSolid = CreateSolidInModel(currentLevel, gData);
                        LevelAndGridSolid secData = new LevelAndGridSolid(levSolid, currentLevel, gData);
                        preResult.Add(secData);
                    }
                }
            }

            // Очистка от солидов, для вспомогательных уровней внутри секций
            foreach (LevelAndGridSolid secData in preResult)
            {
                LevelAndGridSolid[] currentSectionAndAboveLevelsColl;
                if (secData.CurrentLevelData.CurrentAboveLevel != null)
                {
                    currentSectionAndAboveLevelsColl = preResult
                        .Where(r =>
                            r.GridData.CurrentSection.Equals(secData.GridData.CurrentSection)
                            && r.CurrentLevelData.CurrentAboveLevel != null
                            && r.CurrentLevelData.CurrentAboveLevel.Id == secData.CurrentLevelData.CurrentAboveLevel.Id)
                        .ToArray();
                }
                else
                {
                    currentSectionAndAboveLevelsColl = preResult
                        .Where(r =>
                            r.GridData.CurrentSection.Equals(secData.GridData.CurrentSection)
                            && r.CurrentLevelData.CurrentAboveLevel == null)
                        .ToArray();
                }

                if (currentSectionAndAboveLevelsColl.Count() == 1)
                    result.Add(secData);
                else if (currentSectionAndAboveLevelsColl.Any())
                {
                    LevelAndGridSolid minSecData = currentSectionAndAboveLevelsColl
                        .Aggregate((lvlMinElv, x) =>
                            (x.CurrentLevelData.CurrentLevel.Elevation <
                             lvlMinElv.CurrentLevelData.CurrentLevel.Elevation)
                                ? x
                                : lvlMinElv);
                    if (!result.Contains(minSecData))
                        result.Add(minSecData);
                }
                else
                {
                    result.Add(secData);
                }
            }

            // Проверка на коллизии между Solid
            foreach (LevelAndGridSolid secData1 in result)
            {
                foreach (LevelAndGridSolid secData2 in result)
                {
                    // Для паркинга допустимы пересечения солидов уровней
                    if (
                        secData1.CurrentLevelData.CurrentSectionNumber.Equals(LevelData.ParLvlName)
                        || secData2.CurrentLevelData.CurrentSectionNumber.Equals(LevelData.ParLvlName)
                        || secData1.CurrentLevelData.CurrentSectionNumber.Equals(LevelData.StilLvlName)
                        || secData2.CurrentLevelData.CurrentSectionNumber.Equals(LevelData.StilLvlName)
                        ) 
                        continue;

                    if (secData1.Equals(secData2)) 
                        continue;
                    
                    Solid intersectionSolid = BooleanOperationsUtils
                        .ExecuteBooleanOperation(secData1.CurrentSolid,
                            secData2.CurrentSolid, BooleanOperationsType.Intersect);
                    if (intersectionSolid != null && intersectionSolid.Volume > 0)
                        throw new CheckerException(
                            "Солиды уровней пересекаются (ошибка в заполнении параметров сепарации объекта, либо уровни названы не по BEP). " +
                            "Отправь разработчику: " +
                            $"Уровни id: {secData1.CurrentLevelData.CurrentLevel.Id}, " +
                            $"{secData2.CurrentLevelData.CurrentLevel.Id} " +
                            $"для секции №{secData1.GridData.CurrentSection} и " +
                            $"для секции №{secData2.GridData.CurrentSection}");
                }
            }

            return result;
        }

        /// <summary>
        /// Создание Solid по параметрам
        /// </summary>
        /// <param name="levData">Данные по уровням</param>
        /// <param name="grData">Данные по осям</param>
        /// <returns>Созданная по солиду геометрия</returns>
        private static Solid CreateSolidInModel(LevelData levData, GridData grData)
        {
            List<XYZ> pointsOfGridsIntersect = GetPointsOfGridsIntersection(grData.CurrentGrids);
            pointsOfGridsIntersect.Sort(new PntComparer(GetCenterPointOfPoints(pointsOfGridsIntersect)));

            List<XYZ> pointsOfGridsIntersectDwn = new List<XYZ>();
            List<XYZ> pointsOfGridsIntersectUp = new List<XYZ>();
            foreach (XYZ point in pointsOfGridsIntersect)
            {
                XYZ newPointDwn = new XYZ(point.X, point.Y, levData.MinAndMaxLvlPnts[0]);
                pointsOfGridsIntersectDwn.Add(newPointDwn);
                XYZ newPointUp = new XYZ(point.X, point.Y, levData.MinAndMaxLvlPnts[1]);
                pointsOfGridsIntersectUp.Add(newPointUp);
            }

            List<Curve> curvesListDwn = GetCurvesListFromPoints(pointsOfGridsIntersectDwn);
            List<Curve> curvesListUp = GetCurvesListFromPoints(pointsOfGridsIntersectUp);
            CurveLoop curveLoopDwn = CurveLoop.Create(curvesListDwn);
            CurveLoop curveLoopUp = CurveLoop.Create(curvesListUp);
            try
            {
                CurveLoop[] curves = new CurveLoop[] { curveLoopDwn, curveLoopUp };
                SolidOptions solidOptions = new SolidOptions(ElementId.InvalidElementId, ElementId.InvalidElementId);
                return GeometryCreationUtilities.CreateLoftGeometry(curves, solidOptions);
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException)
            {
                StringBuilder sb = new StringBuilder();
                foreach (Grid gr in grData.CurrentGrids)
                {
                    sb.Append(gr.Id);
                    sb.Append(", ");
                }

                throw new CheckerException(
                    "Пограничные оси обязательно должны пересекаться! Проверь оси id: " +
                    $"{sb.ToString().TrimEnd(", ".ToArray())}");
            }
        }

        /// <summary>
        /// Точки пересечения осей
        /// </summary>
        /// <param name="grids">Список осей</param>
        private static List<XYZ> GetPointsOfGridsIntersection(HashSet<Grid> grids)
        {
            List<XYZ> pointsOfGridsIntersect = new List<XYZ>();
            foreach (Grid grid1 in grids)
            {
                if (grid1 == null) 
                    continue;
                
                Curve curve1 = grid1.Curve;
                foreach (Grid grid2 in grids)
                {
                    if (grid2 == null) 
                        continue;
                    
                    if (grid1.Id == grid2.Id) 
                        continue;
                    
                    Curve curve2 = grid2.Curve;
                    curve1.Intersect(curve2, out IntersectionResultArray intersectionResultArray);

                    // Линии не пересекаются. Нужно проверить векторами
                    if (intersectionResultArray == null || intersectionResultArray.IsEmpty)
                    {
                        XYZ vectorIntersection = GetVectorsIntersectPnt(
                            curve1.GetEndPoint(0), 
                            curve1.GetEndPoint(1), 
                            curve2.GetEndPoint(0), 
                            curve2.GetEndPoint(1));

                        if (vectorIntersection != null
                            && !pointsOfGridsIntersect.Any(pgi => vectorIntersection.IsAlmostEqualTo(pgi)))
                        {
                            pointsOfGridsIntersect.Add(vectorIntersection);
                        }
                    }
                    // Линии пересекаются, получаем результат
                    else
                    {
                        foreach (IntersectionResult intersection in intersectionResultArray)
                        {
                            XYZ point = intersection.XYZPoint;
                            if (!pointsOfGridsIntersect.Any(pgi => point.IsAlmostEqualTo(pgi)))
                                pointsOfGridsIntersect.Add(point);
                        }
                    }
                }
            }
            
            return pointsOfGridsIntersect;
        }

        private static XYZ GetVectorsIntersectPnt(XYZ startPoint1, XYZ endPoint1, XYZ startPoint2, XYZ endPoint2)
        {
            XYZ direction1 = endPoint1 - startPoint1;
            XYZ direction2 = endPoint2 - startPoint2;

            // Проверяем на параллельность (если векторное произведение = 0, то линии параллельны)
            XYZ crossProduct = direction1.CrossProduct(direction2);
            
            // Линии параллельны и не пересекаются
            if (crossProduct.IsZeroLength())
                return null;
            // Линии не параллельны, проверим их на пересечение
            else
            {
                // Рассчитываем параметр t1 для пересечения первой линии с продолжением второй линии
                double t1 = ((startPoint2 - startPoint1).CrossProduct(direction2)).DotProduct(crossProduct) 
                            / crossProduct.DotProduct(crossProduct);

                // Рассчитываем параметр t2 для пересечения второй линии с продолжением первой линии
                double t2 = ((startPoint2 - startPoint1).CrossProduct(direction1)).DotProduct(crossProduct) 
                            / crossProduct.DotProduct(crossProduct);

                // Анализ точки пересечения на избыточную удаленность (например, удалено больше чем на 30 м)
                XYZ tempIntersectionPnt = startPoint1 + t1 * direction1;
                double distance = Math.Abs(startPoint1.DistanceTo(tempIntersectionPnt)) - Math.Abs(startPoint1.DistanceTo(endPoint1));
                if (Math.Abs(distance) > 100 )
                    return null;

                // 𝑡2 даёт точку пересечения на второй линии, но в нашем случае достаточно использовать только t1, чтобы получить ту же самую точку пересечения в мировых координатах.
                // То есть одной точки пересечения достаточно, и мы можем выбрать любую из двух линий для её вычисления.
                return startPoint1 + t1 * direction1;
            }
        }

        /// <summary>
        /// Центр между точками
        /// </summary>
        /// <param name="pointsOfGridsIntersect">Список точек</param>
        private static XYZ GetCenterPointOfPoints(List<XYZ> pointsOfGridsIntersect)
        {
            double totalX = 0, totalY = 0, totalZ = 0;
            
            foreach (XYZ xyz in pointsOfGridsIntersect)
            {
                totalX += xyz.X;
                totalY += xyz.Y;
                totalZ += xyz.Z;
            }
            
            double centerX = totalX / pointsOfGridsIntersect.Count;
            double centerY = totalY / pointsOfGridsIntersect.Count;
            double centerZ = totalZ / pointsOfGridsIntersect.Count;

            return new XYZ(centerX, centerY, centerZ);
        }

        private static List<Curve> GetCurvesListFromPoints(List<XYZ> pointsOfGridsIntersect)
        {
            List<Curve> curvesList = new List<Curve>();
            for (int i = 0; i < pointsOfGridsIntersect.Count; i++)
            {
                if (i == pointsOfGridsIntersect.Count - 1)
                {
                    curvesList.Add(Line.CreateBound(pointsOfGridsIntersect[i], pointsOfGridsIntersect[0]));
                    continue;
                }
                curvesList.Add(Line.CreateBound(pointsOfGridsIntersect[i], pointsOfGridsIntersect[i + 1]));

            }
            return curvesList;
        }
    }
}
