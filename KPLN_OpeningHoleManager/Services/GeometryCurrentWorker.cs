using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_OpeningHoleManager.Core;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_OpeningHoleManager.Services
{
    /// <summary>
    /// Локальная утилита работы с геометрией
    /// </summary>
    internal static class GeometryCurrentWorker
    {
        internal static double GetMinimumDistanceBetweenSolids(Solid solid1, Solid solid2)
        {
            double minDistance = double.MaxValue;

            // Прайдзіся па ўсіх гранях першага соліда
            foreach (Face face1 in solid1.Faces)
            {
                Mesh mesh1 = face1.Triangulate();

                foreach (XYZ v1 in mesh1.Vertices)
                {
                    // Для кожнай кропкі на першым солідзе, вызнач бліжэйшую кропку на другім
                    foreach (Face face2 in solid2.Faces)
                    {
                        IntersectionResult intResult = face2.Project(v1);
                        if (intResult != null)
                        {
                            // Праектуем кропку на паверхню другога соліда
                            XYZ projected = intResult.XYZPoint;
                            double distance = v1.DistanceTo(projected);

                            if (distance < minDistance)
                                minDistance = distance;
                        }
                    }
                }
            }

            return minDistance;
        }

        /// <summary>
        /// Получить общий BoundingBoxXYZ на основе коллекции либо элементов Revit, либо IOSOpeningHoleTaskEntity
        /// </summary>
        internal static BoundingBoxXYZ CreateOverallBBox(object[] objColl)
        {
            BoundingBoxXYZ resultBBox = null;

            foreach (object objElem in objColl)
            {
                BoundingBoxXYZ elemBox;
                Transform trans;
                if (objElem is Element elem)
                {
                    elemBox = elem.get_BoundingBox(null);
                    trans = elemBox.Transform;

                    if (elemBox == null)
                    {
                        HtmlOutput.Print($"Ошибка анализа, могут быть не предвиденные результат (проверь отдельно)." +
                            $" У элемента с id:{elem.Id} из вашей модели - нет BoundingBoxXYZ",
                            MessageType.Error);
                        continue;
                    }
                }
                else if (objElem is IOSOpeningHoleTaskEntity iosTask)
                {
                    elemBox = iosTask.IEDElem.get_BoundingBox(null);
                    trans = iosTask.OHE_LinkTransform;

                    if (elemBox == null)
                    {
                        HtmlOutput.Print($"Ошибка анализа, могут быть не предвиденные результат (проверь отдельно)." +
                            $" У элемента с id:{iosTask.IEDElem.Id} из связи {iosTask.OHE_LinkDocument.Title} - нет BoundingBoxXYZ",
                            MessageType.Error);
                        continue;
                    }
                }
                else
                    throw new Exception($"Ошибка - в метод 'CreateOverallBBox' подан тип, который не подвергается анализу");

                if (resultBBox == null)
                {
                    resultBBox = new BoundingBoxXYZ
                    {
                        Min = trans.OfPoint(elemBox.Min),
                        Max = trans.OfPoint(elemBox.Max)
                    };
                }
                else
                {
                    resultBBox.Min = trans.OfPoint(new XYZ(
                        Math.Min(resultBBox.Min.X, elemBox.Min.X),
                        Math.Min(resultBBox.Min.Y, elemBox.Min.Y),
                        Math.Min(resultBBox.Min.Z, elemBox.Min.Z)));

                    resultBBox.Max = trans.OfPoint(new XYZ(
                        Math.Max(resultBBox.Max.X, elemBox.Max.X),
                        Math.Max(resultBBox.Max.Y, elemBox.Max.Y),
                        Math.Max(resultBBox.Max.Z, elemBox.Max.Z)));
                }
            }

            return resultBBox;
        }

        /// <summary>
        /// Создать Outline по указанному BoundingBoxXYZ и расширению
        /// </summary>
        internal static Outline CreateOutline_ByBBoxANDExpand(BoundingBoxXYZ bbox, XYZ expandXYZ)
        {
            Outline resultOutlie;

            Transform bboxTrans = bbox.Transform;

            XYZ bboxMin = bboxTrans.OfPoint(bbox.Min);
            XYZ bboxMax = bboxTrans.OfPoint(bbox.Max);

            // Подготовка расширенного BoundingBoxXYZ, чтобы не упустить эл-ты
            BoundingBoxXYZ expandedCropBB = new BoundingBoxXYZ()
            {
                Max = bboxMax + expandXYZ,
                Min = bboxMin - expandXYZ,
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
        /// Получить вектор для основы отверстия
        /// </summary>
        internal static XYZ GetHostDirection(Element host)
        {
            // Получаю вектор для стены
            XYZ wallDirection;
            if (host is Wall wall)
            {
                Curve curve = (wall.Location as LocationCurve).Curve ??
                    throw new Exception($"Не обработанная основа (не Curve) с id: {host.Id}. Отправь разработчику!");

                LocationCurve wallLocCurve = wall.Location as LocationCurve;
                if (wallLocCurve.Curve is Line wallLine)
                    wallDirection = wallLine.Direction;
                else
                    throw new Exception($"Пока умею работать только с прямыми стенами :(");


            }
            else
                throw new Exception($"Не обработанная основа с id: {host.Id}. Отправь разработчику!");


            return wallDirection;
        }


        /// <summary>
        /// Получить плоскость у солида с определенным углом к указанному вектору 
        /// </summary>
        internal static Face GetFace_ByAngleToDirection(Solid checkSolid, XYZ hostDirection, double tolerance = 0)
        {
            Face result = null;

            double tempArea = 0;
            foreach (Face face in checkSolid.Faces)
            {
                XYZ checkOrigin;
                if (face is PlanarFace planarFace)
                {
                    XYZ faceNormal = planarFace.FaceNormal;
                    checkOrigin = new XYZ(faceNormal.X, faceNormal.Y, hostDirection.Z);

                }
                else if (face is RevolvedFace revFace)
                {
                    XYZ faceNormal = revFace.Axis;
                    checkOrigin = new XYZ(faceNormal.X, faceNormal.Y, hostDirection.Z);
                }
                else continue;
#if Debug2020 || Revit2020
                double angle = UnitUtils.ConvertFromInternalUnits(hostDirection.AngleTo(checkOrigin), DisplayUnitType.DUT_DEGREES_AND_MINUTES);
#else
                double angle = UnitUtils.ConvertFromInternalUnits(hostDirection.AngleTo(checkOrigin), new ForgeTypeId("autodesk.unit.unit:degrees-1.0.1"));
#endif
                // По непонятной причине - площадь может быть отрицательной...
                if ((Math.Round(angle, 2) - 90 <= tolerance || Math.Round(angle, 2) - 180 <= tolerance) 
                    && tempArea < Math.Abs(face.Area))
                {
                    result = face;
                    tempArea = face.Area;
                }
            }

            return result;
        }


        /// <summary>
        /// Получить высоту и ширину (именно в таком порядке) для СОЛИДА относительно вектора размещения солида
        /// </summary>
        /// <param name="solid">Солид для анализа</param>
        /// <param name="mainDirection">Вектор, на который проецируем точки</param>
        /// <returns></returns>
        internal static double[] GetSolidWidhtAndHeight_ByDirection(Solid solid, XYZ mainDirection)
        {
            // Атрымліваем дзве артагональныя восі
            XYZ up = XYZ.BasisZ;
            if (Math.Abs(mainDirection.DotProduct(up)) > 0.99)
                up = XYZ.BasisX; // калі mainDirection амаль Z, то выбіраем іншы
            XYZ sideDirection = mainDirection.CrossProduct(up);
            up = sideDirection.CrossProduct(mainDirection);

            // Збіраем усе вяршыні
            List<XYZ> allPoints = new List<XYZ>();
            foreach (Face face in solid.Faces)
            {
                foreach (EdgeArray edgeArray in face.EdgeLoops)
                {
                    foreach (Edge edge in edgeArray)
                    {
                        allPoints.AddRange(edge.Tessellate());
                    }
                }
            }

            // Праектуем усе пункты ў мясцовую сістэму каардынат
            List<XYZ> projected = allPoints
                .Select(p =>
                    {
                        double x = p.DotProduct(mainDirection);
                        double y = p.DotProduct(sideDirection);
                        double z = p.DotProduct(up);
                        return new XYZ(x, y, z);
                    })
                .ToList();

            // Знаходзім памеры
            double minX = projected.Min(p => p.X);
            double maxX = projected.Max(p => p.X);
            double minY = projected.Min(p => p.Y);
            double maxY = projected.Max(p => p.Y);
            double minZ = projected.Min(p => p.Z);
            double maxZ = projected.Max(p => p.Z);

            double width = maxX - minX;
            double height = maxZ - minZ;

            double[] result = new double[2]
            {
                width,
                height
            };

            return result;
        }
    }
}
