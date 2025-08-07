using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Parameters_Ribbon.Services
{
    internal static class GeometryWorker
    {
        /// <summary>
        /// Получить сумарный Solid для элемента
        /// </summary>
        internal static Solid GetRevitElemUniontSolid(Element elem, Transform transform = null)
        {
            Options opt = new Options
            {
                DetailLevel = ViewDetailLevel.Fine,
                ComputeReferences = true
            };

            GeometryElement geomElem = elem.get_Geometry(opt);
            if (geomElem != null)
            {
                // Получаю список солидов
                List<Solid> solidColl = new List<Solid>();
                GetSolidsFromGeomElem(geomElem, Transform.Identity, solidColl);

                // Если солидов нет - создаю свой солид по BoundingBoxXYZ
                if (solidColl.Count() == 0)
                {
                    BoundingBoxXYZ bbox = elem.get_BoundingBox(null);
                    XYZ bboxMin = bbox.Min;
                    XYZ bboxMax = bbox.Max;

                    #region Уточнение координат по плоским эл-м
                    // Очистка от плоских эл-в (по оси X)
                    if (Math.Round(bboxMin.X, 5) == Math.Round(bboxMax.X, 5))
                    {
                        double newMinX = bboxMin.X > 0 ? bboxMin.X - 0.1 : bboxMin.X + 0.1;
                        double newMaxX = bboxMax.X > 0 ? bboxMax.X + 0.1 : bboxMax.X - 0.1;
                        bboxMin = new XYZ(newMinX, bboxMin.Y, bboxMin.Z);
                        bboxMax = new XYZ(newMaxX, bboxMax.Y, bboxMax.Z);
                    }
                    // Очистка от плоских эл-в (по оси Y)
                    if (Math.Round(bboxMin.Y, 5) == Math.Round(bboxMax.Y, 5))
                    {
                        double newMinY = bboxMin.Y > 0 ? bboxMin.Y - 0.1 : bboxMin.Y + 0.1;
                        double newMaxY = bboxMax.Y > 0 ? bboxMax.Y + 0.1 : bboxMax.Y - 0.1;
                        bboxMin = new XYZ(bboxMin.X, newMinY, bboxMin.Z);
                        bboxMax = new XYZ(bboxMax.X, newMaxY, bboxMax.Z);
                    }
                    // Очистка от плоских эл-в (по оси Z)
                    if (Math.Round(bboxMin.Z, 5) == Math.Round(bboxMax.Z, 5))
                        bboxMax = new XYZ(bboxMax.X, bboxMax.Y, bboxMin.Z + 0.1);
                    #endregion

                    List<XYZ> pointsDwn = new List<XYZ>()
                    {
                        new XYZ(bboxMin.X, bboxMin.Y, bboxMin.Z),
                        new XYZ(bboxMin.X, bboxMax.Y, bboxMin.Z),
                        new XYZ(bboxMax.X, bboxMax.Y, bboxMin.Z),
                        new XYZ(bboxMax.X, bboxMin.Y, bboxMin.Z),
                    };
                    List<XYZ> pointsUp = new List<XYZ>()
                    {
                        new XYZ(bboxMin.X, bboxMin.Y, bboxMax.Z),
                        new XYZ(bboxMin.X, bboxMax.Y, bboxMax.Z),
                        new XYZ(bboxMax.X, bboxMax.Y, bboxMax.Z),
                        new XYZ(bboxMax.X, bboxMin.Y, bboxMax.Z),
                    };


                    List<Curve> curvesListDwn = GetCurvesListFromPoints(pointsDwn);
                    List<Curve> curvesListUp = GetCurvesListFromPoints(pointsUp);
                    CurveLoop curveLoopDwn = CurveLoop.Create(curvesListDwn);
                    CurveLoop curveLoopUp = CurveLoop.Create(curvesListUp);

                    CurveLoop[] curves = new CurveLoop[] { curveLoopDwn, curveLoopUp };
                    SolidOptions solidOptions = new SolidOptions(ElementId.InvalidElementId, ElementId.InvalidElementId);
                    solidColl.Add(GeometryCreationUtilities.CreateLoftGeometry(curves, solidOptions));
                }

                if (solidColl.Count() == 0)
                    throw new Exception($"У элемента с именем \"{elem.Name}\" с id: {elem.Id} из модели {elem.Document.Title} проблемы с получением Solid (нет солидов с объёмом). " +
                        $"Отправь разработчику.");

                // Объединяю солиды в резултирующий
                Solid unionSolid = null;
                foreach (Solid solid in solidColl)
                {
                    try
                    {
                        if (unionSolid == null)
                        {
                            unionSolid = solid;
                            continue;
                        }
                    
                        unionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(unionSolid, solid, BooleanOperationsType.Union);
                    }
                    // Могут быть проблемы с тем, что нельзя выполнить операцию. Игнорим, остальных солидов должно хватить
                    catch (Autodesk.Revit.Exceptions.InvalidOperationException) { continue; }
                    catch (Exception ex) { throw ex; }
                }

                if (unionSolid == null)
                    throw new Exception($"У элемента с именем \"{elem.Name}\" с id: {elem.Id} из модели {elem.Document.Title} проблемы с получением Solid (геометрия не объединяется). " +
                        $"Отправь разработчику.");

                // Трансформ по координатам (если нужно)
                if (transform != null)
                    unionSolid = SolidUtils.CreateTransformed(unionSolid, transform);

                return unionSolid;
            }

            return null;
        }

        /// <summary>
        /// Получить BoundingBoxXYZ из Solid
        /// </summary>
        internal static BoundingBoxXYZ GetBoundingBoxXYZ(Solid solid, Transform trans = null)
        {
            BoundingBoxXYZ bbox = solid.GetBoundingBox();
            if (trans == null)
                trans = bbox.Transform;


            return new BoundingBoxXYZ()
            {
                Max = trans.OfPoint(bbox.Max),
                Min = trans.OfPoint(bbox.Min),
            };
        }

        /// <summary>
        /// Создать Outline по указанному BoundingBoxXYZ и расширению
        /// </summary>
        internal static Outline CreateOutline_ByBBoxANDExpand(BoundingBoxXYZ bbox, double expandValue)
        {
            Outline resultOutlie;

            Transform bboxTrans = bbox.Transform;

            XYZ bboxMin = bboxTrans.OfPoint(bbox.Min);
            XYZ bboxMax = bboxTrans.OfPoint(bbox.Max);

            // Подготовка расширенного BoundingBoxXYZ, чтобы не упустить эл-ты
            BoundingBoxXYZ expandedCropBB = new BoundingBoxXYZ()
            {
                Max = bboxMax + new XYZ(expandValue, expandValue, expandValue),
                Min = bboxMin - new XYZ(expandValue, expandValue, expandValue),
            };

            resultOutlie = new Outline(expandedCropBB.Min, expandedCropBB.Max);

            if (resultOutlie.IsEmpty)
            {
                XYZ transExpandedElemBBMin = resultOutlie.MinimumPoint;
                XYZ transExpandedElemBBMax = resultOutlie.MaximumPoint;

                double minX = transExpandedElemBBMin.X;
                double minY = transExpandedElemBBMin.Y;
                double minZ = transExpandedElemBBMin.Z;

                double maxX = transExpandedElemBBMax.X;
                double maxY = transExpandedElemBBMax.Y;
                double maxZ = transExpandedElemBBMax.Z;

                double sminX = Math.Min(minX, maxX);
                double sminY = Math.Min(minY, maxY);
                double sminZ = Math.Min(minZ, maxZ);

                double smaxX = Math.Max(minX, maxX);
                double smaxY = Math.Max(minY, maxY);
                double smaxZ = Math.Max(minZ, maxZ);

                XYZ pntMin = new XYZ(sminX, sminY, sminZ);
                XYZ pntMax = new XYZ(smaxX, smaxY, smaxZ);

                resultOutlie = new Outline(pntMin, pntMax);
            }

            if (!resultOutlie.IsEmpty && resultOutlie.IsValidObject)
                return resultOutlie;

            throw new Exception($"Отправь разработчику - не удалось создать Outline для фильтрации");
        }


        /// <summary>
        /// Получить Solidы из элемента
        /// </summary>
        private static void GetSolidsFromGeomElem(GeometryElement geometryElement, Transform transformation, IList<Solid> solidColl)
        {
            foreach (GeometryObject geomObject in geometryElement)
            {
                switch (geomObject)
                {
                    case Solid solid:
                        // Очень маленькие солиды - мусор, и могут выдать ошибку при объединении
                        if (solid.Volume > 0.05)
                            solidColl.Add(solid);

                        break;

                    case GeometryInstance geomInstance:
                        GetSolidsFromGeomElem(geomInstance.GetInstanceGeometry(), geomInstance.Transform.Multiply(transformation), solidColl);
                        break;

                    case GeometryElement geomElem:
                        GetSolidsFromGeomElem(geomElem, transformation, solidColl);
                        break;
                }
            }
        }

        /// <summary>
        /// Создать кривые по точкам пересечения
        /// </summary>
        private static List<Curve> GetCurvesListFromPoints(List<XYZ> pointsOfIntersect)
        {
            List<Curve> curvesList = new List<Curve>();
            for (int i = 0; i < pointsOfIntersect.Count; i++)
            {
                if (i == pointsOfIntersect.Count - 1)
                {
                    curvesList.Add(Line.CreateBound(pointsOfIntersect[i], pointsOfIntersect[0]));
                    continue;
                }

                curvesList.Add(Line.CreateBound(pointsOfIntersect[i], pointsOfIntersect[i + 1]));
            }
            return curvesList;
        }
    }
}
