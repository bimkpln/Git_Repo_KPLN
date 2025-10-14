using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace KPLN_ModelChecker_Lib.Commands.Common.CheckMEPHeight
{
    internal static class GeometryUtil
    {
        /// <summary>
        /// Установить коллекцию солидов элемента в коллецию
        /// </summary>
        /// <param name="geometryElement">Эл-т геометрии для получения солида</param>
        /// <param name="transformation">Трансформ по координатам</param>
        /// <param name="solids">Коллекция для изменения</param>
        internal static void SetSolidsFromGeomElem(GeometryElement geometryElement, Transform transformation, List<Solid> solids)
        {
            foreach (GeometryObject geomObject in geometryElement)
            {
                switch (geomObject)
                {
                    case Solid solid:
                        if (solid.Volume > 0) solids.Add(solid);
                        break;

                    case GeometryInstance geomInstance:
                        SetSolidsFromGeomElem(geomInstance.GetInstanceGeometry(), geomInstance.Transform.Multiply(transformation), solids);
                        break;

                    case GeometryElement geomElem:
                        SetSolidsFromGeomElem(geomElem, transformation, solids);
                        break;
                }
            }
        }

        /// <summary>
        /// Получить BoundingBoxXYZ
        /// </summary>
        /// <param name="geomElem">Эл-т геометрии для получения BoundingBoxXYZ</param>
        /// <param name="transform">Трансформ для мутации</param>
        /// <returns></returns>
        internal static BoundingBoxXYZ GetBoundingBoxXYZ(GeometryElement geomElem, Transform transform = null)
        {
            foreach (GeometryObject obj in geomElem)
            {
                Solid solid = obj as Solid;
                return GetBoundingBoxXYZ(solid, transform);
            }

            return null;
        }

        /// <summary>
        /// Получить BoundingBoxXYZ
        /// </summary>
        /// <param name="solid">Солид для анализа</param>
        /// <param name="transform">Трансформ для мутации</param>
        /// <returns></returns>
        internal static BoundingBoxXYZ GetBoundingBoxXYZ(Solid solid, Transform transform = null)
        {
            if (solid != null && solid.Volume != 0)
            {
                BoundingBoxXYZ bbox = solid.GetBoundingBox();
                Transform bboxTrans = bbox.Transform;
                
                Transform resultTrans;
                if (transform != null)
                    resultTrans = bboxTrans * transform;
                else
                    resultTrans = bboxTrans;

                return new BoundingBoxXYZ()
                {
                    Max = resultTrans.OfPoint(bbox.Max),
                    Min = resultTrans.OfPoint(bbox.Min),
                };
            }

            return null;
        }

        
    }
}
