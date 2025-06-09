﻿using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
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
        private static LogicalOrFilter _elemCatLogicalOrFilter;
        private static Func<Element, bool> _elemFilterFunc;

        /// <summary>
        /// Фильтр для ключевой фильтрации по категориям
        /// </summary>
        internal static LogicalOrFilter ElemCatLogicalOrFilter
        {
            get
            {
                if (_elemCatLogicalOrFilter == null)
                {
                    List<ElementFilter> catFilters = new List<ElementFilter>();
                    catFilters.AddRange(BuiltInCategories.Select(bic => new ElementCategoryFilter(bic)));

                    _elemCatLogicalOrFilter = new LogicalOrFilter(catFilters);
                }

                return _elemCatLogicalOrFilter;
            }
        }

        /// <summary>
        /// Общая функция для фильтра для ДОПОЛНИЕТЛЬНОГО просеивания элементов модели (новых и отредактированных)
        /// </summary>
        internal static Func<Element, bool> ElemExtraFilterFunc
        {
            get
            {
                if (_elemFilterFunc == null)
                {
                    _elemFilterFunc = (el) =>
                        el.Category != null
                        && !(el is ElementType)
                        // Молниезащита ЭОМ
                        && !(el.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().StartsWith("Полоса_")
                            || el.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().StartsWith("Пруток_")
                            || el.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().StartsWith("Уголок_")
                            // Фильтрация семейств без геометрии от Ostec, крышка лотка DKC, неподвижную опору ОВВК
                            || (el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().ToLower().Contains("ostec") && (el is FamilyInstance fi && fi.SuperComponent != null))
                            || el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().ToLower().Contains("470_dkc_s5_accessories")
                            || el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().ToLower().Contains("757_опора_неподвижная")
                            // Фильтрация семейств под которое НИКОГДА не должно быть отверстий
                            || el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().StartsWith("501_")
                            || el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().StartsWith("551_")
                            || el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().StartsWith("556_")
                            || el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().StartsWith("557_")
                            || el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().StartsWith("560_")
                            || el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().StartsWith("561_")
                            || el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().StartsWith("565_")
                            || el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().StartsWith("570_")
                            || el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().StartsWith("582_")
                            || el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().StartsWith("592_"));
                }

                return _elemFilterFunc;
            }
        }


        /// <summary>
        /// Список BuiltInCategory для файлов ИОС, которые обрабатываются
        /// </summary>
        private static List<BuiltInCategory> BuiltInCategories
        {
            get =>
                new List<BuiltInCategory>()
                { 
                    // ОВВК (ЭОМСС - огнезащита)
                    BuiltInCategory.OST_DuctCurves,
                    BuiltInCategory.OST_DuctFitting,
                    BuiltInCategory.OST_DuctTerminal,
                    BuiltInCategory.OST_DuctAccessory,
                    BuiltInCategory.OST_PipeCurves,
                    BuiltInCategory.OST_PipeFitting,
                    BuiltInCategory.OST_PipeAccessory,
                    BuiltInCategory.OST_MechanicalEquipment,
                    // ЭОМСС
                    BuiltInCategory.OST_CableTray,
                    BuiltInCategory.OST_CableTrayFitting,
                };
        }

        /// <summary>
        /// Установить коллекцию IOSElemEntities для ARKRElemEntity ПО ВЫБРАННЫМ ЭЛ-ТАМ
        /// </summary>
        /// <returns></returns>
        internal static ARKRElemEntity SetIOSEntities_BySelectedIOSElems(ARKRElemEntity arkrEntity, RevitLinkInstance iosLinkInst, Element iosLinkElem)
        {
            // Подготовка линка к обработке
            Document iosLinkDoc = iosLinkInst.GetLinkDocument();
            Transform iosLinkTransfrom = GetLinkTransform(iosLinkInst);

            IOSElemEntity iosEnt = GetIOSElemEntities_BySolidIntersect(arkrEntity, iosLinkDoc, iosLinkElem, iosLinkTransfrom);
            if (iosEnt != null)
                arkrEntity.IOSElemEntities.Add(iosEnt);

            return arkrEntity;
        }

        /// <summary>
        /// Установить коллекцию IOSElemEntities для ARKRElemEntity ПО ВСЕМ ФАЙЛАМ
        /// </summary>
        /// <returns></returns>
        internal static ARKRElemEntity SetIOSEntities_ByLinks(ARKRElemEntity arkrEntity, List<RevitLinkInstance> iosLinkInsts)
        {
            foreach (var iosLinkInst in iosLinkInsts)
            {
                // Подготовка линка к обработке
                Document iosLinkDoc = iosLinkInst.GetLinkDocument();
                Transform iosLinkTransfrom = GetLinkTransform(iosLinkInst);


                // Генерация Outline для быстрого поиска (QuickFilter)
                Outline filterOutline = CreateFilterOutline_BySolid(arkrEntity.ARKRHost_Solid, 2);


                // Генерация потенциальных эл-в (QuickFilter)
                Outline checkOutline = new Outline(iosLinkTransfrom.Inverse.OfPoint(filterOutline.MinimumPoint), iosLinkTransfrom.Inverse.OfPoint(filterOutline.MaximumPoint));
                HashSet<Element> checkLinkElems = new HashSet<Element>(new ElementComparerById());

                BoundingBoxIntersectsFilter intersectsFilter = new BoundingBoxIntersectsFilter(checkOutline, 0.1);
                BoundingBoxIsInsideFilter insideFilter = new BoundingBoxIsInsideFilter(checkOutline, 0.1);

                checkLinkElems.UnionWith(new FilteredElementCollector(iosLinkDoc)
                    .WherePasses(new LogicalAndFilter(ElemCatLogicalOrFilter, intersectsFilter))
                    .Where(ElemExtraFilterFunc));
                checkLinkElems.UnionWith(new FilteredElementCollector(iosLinkDoc)
                    .WherePasses(new LogicalAndFilter(ElemCatLogicalOrFilter, insideFilter))
                    .Where(ElemExtraFilterFunc));


                // Уточнение по пересечениям с хостом из АР (SlowFilter)
                foreach (Element iosLinkElem in checkLinkElems)
                {
                    IOSElemEntity iosEnt = GetIOSElemEntities_BySolidIntersect(arkrEntity, iosLinkDoc, iosLinkElem, iosLinkTransfrom);
                    if (iosEnt != null)
                        arkrEntity.IOSElemEntities.Add(iosEnt);
                }
            }

            return arkrEntity;
        }

        private static IOSElemEntity GetIOSElemEntities_BySolidIntersect(ARKRElemEntity entity, Document iosLinkDoc, Element iosLinkElem, Transform iosLinkTransfrom)
        {
            Solid iosElemSolid = GeometryWorker.GetRevitElemSolid(iosLinkElem, iosLinkTransfrom);
            if (iosElemSolid != null)
            {
                try
                {
                    Solid intersectSolid = BooleanOperationsUtils.ExecuteBooleanOperation(entity.ARKRHost_Solid, iosElemSolid, BooleanOperationsType.Intersect);
                    if (intersectSolid != null && intersectSolid.Volume > 0)
                        return new IOSElemEntity(iosLinkDoc, iosLinkTransfrom, iosLinkElem, iosElemSolid, intersectSolid);
                }
                // Может падать ошибка если тела неточно расположены между собой: https://www.revitapidocs.com/2023/89cb7975-cc76-65ba-b996-bcb78d12161a.htm
                catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                {
                    throw new Exception($"Для эл-та с id: {iosLinkElem.Id} из файла {iosLinkElem.Document.Title} не удалось провести анализ на пересечение со стеной АР с id: {entity.ARKRHost_Element.Id}");
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

            return null;
        }

        /// <summary>
        /// Создать Outline по элементу АР и расширению
        /// </summary>
        private static Outline CreateFilterOutline_BySolid(Solid arHostSolid, double expandValue)
        {
            Outline resultOutlie;

            BoundingBoxXYZ bbox = arHostSolid.GetBoundingBox();
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
        /// Задать Transform для линка
        /// </summary>
        /// <param name="inst">Instance связи</param>
        /// <returns></returns>
        private static Transform GetLinkTransform(RevitLinkInstance linkInst)
        {
            Instance inst = linkInst as Instance;
            Transform instTrans = inst.GetTransform();
            // Метка того, что базис трансформа тождество. Если нет, то создаём такой трансформ
            if (instTrans.IsTranslation)
                return instTrans;
            else
                return Transform.CreateTranslation(instTrans.Origin);
        }
    }
}
