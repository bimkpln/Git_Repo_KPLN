using Autodesk.Revit.DB;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_OpeningHoleManager.Core;
using System;
using System.Collections.Generic;

namespace KPLN_OpeningHoleManager.Services
{
    internal sealed class GeometryWorker
    {
        /// <summary>
        /// Получить SOLID элемента Ревит
        /// </summary>
        public static Solid GetElemSolid(Element elem, Transform transform = null)
        {
            Solid resultSolid = null;
            GeometryElement geomElem = elem.get_Geometry(new Options { DetailLevel = ViewDetailLevel.Fine });
            foreach (GeometryObject gObj in geomElem)
            {
                Solid solid = gObj as Solid;
                GeometryInstance gInst = gObj as GeometryInstance;
                if (solid != null)
                    resultSolid = solid;
                else if (gInst != null)
                {
                    GeometryElement instGeomElem = gInst.GetInstanceGeometry();
                    double tempVolume = 0;
                    foreach (GeometryObject gObj2 in instGeomElem)
                    {
                        solid = gObj2 as Solid;
                        if (solid != null && solid.Volume > tempVolume)
                        {
                            tempVolume = solid.Volume;
                            resultSolid = solid;
                        }
                    }
                }
            }

            if (resultSolid == null && (elem is FamilyInstance fi && fi.SuperComponent == null))
                throw new Exception($"У элемента с id: {elem.Id} из модели {elem.Document.Title} проблемы с получением Solid. Отправь разработчику.");

            // Трансформ по координатам (если нужно)
            if (resultSolid != null && transform != null)
                resultSolid = SolidUtils.CreateTransformed(resultSolid, transform);

            return resultSolid;
        }

        /// <summary>
        /// Получить общий BoundingBoxXYZ на основе всех заданий от ИОС с УЧЁТОМ вшитого Transform
        /// </summary>
        /// <returns></returns>
        public static BoundingBoxXYZ CreateOverallBBox(IOSOpeningHoleTaskEntity[] iosTasks)
        {
            BoundingBoxXYZ resultBBox = null;

            foreach (IOSOpeningHoleTaskEntity iosTask in iosTasks)
            {
                BoundingBoxXYZ elementBox = iosTask.OHE_Element.get_BoundingBox(null);
                if (elementBox == null)
                {
                    HtmlOutput.Print($"Ошибка анализа, могут быть не предвиденные результат (проверь отдельно)." +
                        $" У элемента с id:{iosTask.OHE_Element.Id} из связи {iosTask.OHE_LinkDocument.Title} - нет BoundingBoxXYZ",
                        MessageType.Error);
                    continue;
                }

                if (resultBBox == null)
                {
                    resultBBox = new BoundingBoxXYZ
                    {
                        Min = iosTask.OHE_LinkTransform.OfPoint(elementBox.Min),
                        Max = iosTask.OHE_LinkTransform.OfPoint(elementBox.Max)
                    };
                }
                else
                {
                    resultBBox.Min = iosTask.OHE_LinkTransform.OfPoint(new XYZ(
                        Math.Min(resultBBox.Min.X, elementBox.Min.X),
                        Math.Min(resultBBox.Min.Y, elementBox.Min.Y),
                        Math.Min(resultBBox.Min.Z, elementBox.Min.Z)));

                    resultBBox.Max = iosTask.OHE_LinkTransform.OfPoint(new XYZ(
                        Math.Max(resultBBox.Max.X, elementBox.Max.X),
                        Math.Max(resultBBox.Max.Y, elementBox.Max.Y),
                        Math.Max(resultBBox.Max.Z, elementBox.Max.Z)));
                }
            }

            return resultBBox;
        }

        /// <summary>
        /// Создать Outline по указанному BoundingBoxXYZ и расширению
        /// </summary>
        public static Outline CreateFilterOutline(BoundingBoxXYZ bbox, double expandValue)
        {
            Outline resultOutlie;

            XYZ bboxMin = bbox.Min;
            XYZ bboxMax = bbox.Max;

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

            HtmlOutput.Print($"Отправь разработчику - не удалось создать Outline для фильтрации", MessageType.Error);
            return null;
        }
    }
}
