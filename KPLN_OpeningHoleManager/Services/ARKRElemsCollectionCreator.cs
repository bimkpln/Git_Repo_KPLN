using Autodesk.Revit.DB;
using KPLN_ModelChecker_Lib.Services;
using KPLN_OpeningHoleManager.Core;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_OpeningHoleManager.Services
{
    /// <summary>
    /// Сервис по созданию сущностей ARKRIOSElemEntity
    /// </summary>
    internal static class ARKRElemsCollectionCreator
    {
        /// <summary>
        /// Установить коллекцию IOSElemEntities для ARKRElemEntity ПО ВЫБРАННЫМ ЭЛ-ТАМ
        /// </summary>
        /// <returns></returns>
        internal static ARKRElemEntity SetIOSEntities_BySelectedIOSElems(ARKRElemEntity arkrEntity, RevitLinkInstance iosLinkInst, Element iosLinkElem)
        {
            // Подготовка линка к обработке
            Document iosLinkDoc = iosLinkInst.GetLinkDocument();
            Transform iosLinkTransfrom = GetLinkTransform(iosLinkInst);

            IOSElemEntity iosEnt = GetIOSElemEntity_BySolidIntersect(arkrEntity, iosLinkDoc, iosLinkElem, iosLinkTransfrom);
            if (iosEnt != null)
                arkrEntity.IOSElemEntities.Add(iosEnt);

            return arkrEntity;
        }

        /// <summary>
        /// Установить коллекцию IOSElemEntities для ARKRElemEntity ПО ВСЕМ ФАЙЛАМ
        /// </summary>
        /// <returns></returns>
        internal static ARKRElemEntity SetIOSEntities_ByIOSElemEntities(
            ARKRElemEntity arkrEntity,
            List<RevitLinkInstance> iosLinkInsts,
            IDictionary<RevitLinkInstance, Transform> linkTransforms,
            IDictionary<RevitLinkInstance, ICollection<Element>> prefilteredLinkElems = null)
        {
            // Генерация Outline для быстрого поиска (QuickFilter)
            Outline filterOutline = GeometryCurrentWorker.CreateOutline_ByBBoxANDExpand(arkrEntity.IGDSolid.GetBoundingBox(), new XYZ(0.5, 0.5, 0.5));

            foreach (RevitLinkInstance iosLinkInst in iosLinkInsts)
            {
                Document iosLinkDoc = iosLinkInst.GetLinkDocument();
                if (!linkTransforms.TryGetValue(iosLinkInst, out Transform iosLinkTransform))
                    throw new Exception($"Не предустановлен Transform для связи {iosLinkInst.Name}");

                Outline checkOutline = TransformFilterOutline_ToLink(filterOutline, iosLinkTransform);

                BoundingBoxIntersectsFilter intersectsFilter = new BoundingBoxIntersectsFilter(checkOutline, 0.1);
                BoundingBoxIsInsideFilter insideFilter = new BoundingBoxIsInsideFilter(checkOutline, 0.1);

                HashSet<Element> checkLinkElems = new HashSet<Element>(new ElementComparerById());

                if (prefilteredLinkElems.TryGetValue(iosLinkInst, out ICollection<Element> preElems))
                {
                    checkLinkElems.UnionWith(preElems.Where(e => intersectsFilter.PassesFilter(iosLinkDoc, e.Id)));
                    checkLinkElems.UnionWith(preElems.Where(e => insideFilter.PassesFilter(iosLinkDoc, e.Id)));
                }
                else
                    throw new Exception($"Не предустановлены элементы для связи {iosLinkInst.Name}");

                foreach (Element iosLinkElem in checkLinkElems)
                {
                    IOSElemEntity iosEnt = GetIOSElemEntity_BySolidIntersect(arkrEntity, iosLinkDoc, iosLinkElem, iosLinkTransform);
                    if (iosEnt != null)
                        arkrEntity.IOSElemEntities.Add(iosEnt);
                }
            }

            return arkrEntity;
        }

        /// <summary>
        /// Задать Transform для линка
        /// </summary>
        /// <param name="linkInst">Instance связи</param>
        /// <returns></returns>
        internal static Transform GetLinkTransform(RevitLinkInstance linkInst)
        {
            Instance inst = linkInst as Instance;
            Transform instTrans = inst.GetTotalTransform();
            return instTrans;
        }

        /// <summary>
        /// Создание IOSElemEntity по пересечению (если оно есть)
        /// </summary>
        private static IOSElemEntity GetIOSElemEntity_BySolidIntersect(ARKRElemEntity entity, Document iosLinkDoc, Element iosLinkElem, Transform iosLinkTransfrom)
        {
            // Фильтрация по продольным коллизиям
            Location iosLinkElemLoc = iosLinkElem.Location;
            if (iosLinkElemLoc != null
                && iosLinkElemLoc is LocationCurve iosLinkElemLocCurve
                && iosLinkElemLocCurve.Curve is Line iosLinkElemLine)
            {
                XYZ iosDir = iosLinkElemLine.Direction;
                XYZ hostDir = GeometryCurrentWorker.GetHostDirection(entity.IEDElem);
                XYZ crossProd = iosDir.CrossProduct(hostDir);

                // Отсеиваю вертикальные, под 90° участки и горизонтальные, параллельные хосту (они в 99% ошибки, остальное закроем Navisworks)
                // Могут быть ошибки, на тесте все ок, но звучит не убедительно
                // Концепция вот какая - если CrossProduct по Z стремиться к 0, значит элементы в одно плоскости
                if (Math.Round(crossProd.Z, 5) - 0 == 0)
                    return null;
            }

            // Основной анализ
            Solid iosElemSolid = GeometryWorker.GetRevitElemUniontSolid(iosLinkElem, iosLinkTransfrom);
            if (iosElemSolid != null)
            {
                try
                {
                    Solid intersectSolid = BooleanOperationsUtils.ExecuteBooleanOperation(entity.IGDSolid, iosElemSolid, BooleanOperationsType.Intersect);
                    if (intersectSolid != null && intersectSolid.Volume > 0)
                        return new IOSElemEntity(iosLinkDoc, iosLinkElem, intersectSolid);
                }
                // Может падать ошибка если тела неточно расположены между собой: https://www.revitapidocs.com/2023/89cb7975-cc76-65ba-b996-bcb78d12161a.htm
                catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                {
                    throw new Exception($"Для эл-та с id: {iosLinkElem.Id} из файла {iosLinkElem.Document.Title} не удалось провести анализ на пересечение со стеной АР с id: {entity.IEDElem.Id}");
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

            return null;
        }

        /// <summary>
        /// Трансформ для аутлайна, в координаты линка. Возможно опракидывание координат (например Y у MIN будет больше Y у MAX) - это
        /// не допустимо для создания фильтра
        /// </summary>
        private static Outline TransformFilterOutline_ToLink(Outline outlineToTransform, Transform linkTRansform)
        {
            // Inverse - т.к. возвращаюсь в координаты линка
            Outline transformOutline = new Outline(linkTRansform.Inverse.OfPoint(outlineToTransform.MinimumPoint), linkTRansform.Inverse.OfPoint(outlineToTransform.MaximumPoint));

            XYZ transOutlineMin = transformOutline.MinimumPoint;
            XYZ transOutlineMax = transformOutline.MaximumPoint;

            double minX = transOutlineMin.X;
            double minY = transOutlineMin.Y;
            double minZ = transOutlineMin.Z;

            double maxX = transOutlineMax.X;
            double maxY = transOutlineMax.Y;
            double maxZ = transOutlineMax.Z;

            double sminX = Math.Min(minX, maxX);
            double sminY = Math.Min(minY, maxY);
            double sminZ = Math.Min(minZ, maxZ);

            double smaxX = Math.Max(minX, maxX);
            double smaxY = Math.Max(minY, maxY);
            double smaxZ = Math.Max(minZ, maxZ);

            XYZ pntMin = new XYZ(sminX, sminY, sminZ);
            XYZ pntMax = new XYZ(smaxX, smaxY, smaxZ);

            return new Outline(pntMin, pntMax);
        }
    }
}
