using Autodesk.Revit.DB;
using KPLN_ModelChecker_Lib.Commands.Common.CheckHoles;
using KPLN_ModelChecker_Lib.Common;
using KPLN_ModelChecker_Lib.Core;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_Lib.Commands
{
    public sealed class CheckHoles : AbstrCheck
    {
        /// <summary>
        /// Список BuiltInCategory для файлов ИОС, которые обрабатываются
        /// </summary>
        private readonly List<BuiltInCategory> _builtInCategories = new List<BuiltInCategory>()
        { 
            // ОВВК (ЭОМСС - огнезащита)
            BuiltInCategory.OST_DuctCurves,
            BuiltInCategory.OST_DuctInsulations,
            BuiltInCategory.OST_DuctFitting,
            BuiltInCategory.OST_DuctAccessory,
            BuiltInCategory.OST_DuctTerminal,
            BuiltInCategory.OST_PipeCurves,
            BuiltInCategory.OST_PipeInsulations,
            BuiltInCategory.OST_PipeFitting,
            BuiltInCategory.OST_PipeAccessory,
            BuiltInCategory.OST_MechanicalEquipment,
            // ЭОМСС
            BuiltInCategory.OST_CableTray,
            BuiltInCategory.OST_CableTrayFitting,
        };

        private readonly Func<Element, bool> _elemExtraFilterFunc = (el) =>
        {
            if (el.Category == null || el is ElementType)
                return false;

            string elem_type_param = el.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)?.AsValueString()?.ToLower() ?? "";
            string elem_family_param = el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM)?.AsValueString()?.ToLower() ?? "";

            return !(
                // Молниезащита ЭОМ
                elem_type_param.StartsWith("полоса_")
                || elem_type_param.StartsWith("пруток_")
                || elem_type_param.StartsWith("уголок_")
                || elem_type_param.StartsWith("asml_эг_пруток-катанка")
                || elem_type_param.StartsWith("asml_эг_полоса")
                // Фильтрация семейств без геометрии от Ostec, крышка лотка DKC, неподвижную опору ОВВК
                || (elem_family_param.Contains("ostec") && (el is FamilyInstance fi && fi.SuperComponent != null))
                || elem_family_param.Contains("470_dkc_s5_accessories")
                || elem_family_param.Contains("470_dkc_fireproof_out")
                || elem_family_param.Contains("dkc_ceiling")
                || elem_family_param.Contains("757_опора_неподвижная")
                // Фильтрация семейств под которое НИКОГДА не должно быть отверстий
                || elem_family_param.StartsWith("501_")
                || elem_family_param.StartsWith("551_")
                || elem_family_param.StartsWith("552_")
                || elem_family_param.StartsWith("556_")
                || elem_family_param.StartsWith("557_")
                || elem_family_param.StartsWith("560_")
                || elem_family_param.StartsWith("561_")
                || elem_family_param.StartsWith("565_")
                || elem_family_param.StartsWith("570_")
                || elem_family_param.StartsWith("582_")
                || elem_family_param.StartsWith("592_")
                // Фильтрация типов семейств для которых опытным путём определено, что солид у них не взять (очень сложные семейства)
                || elem_type_param.Contains("узел учета квартиры для гвс")
                || elem_type_param.Contains("узел учета офиса для гвс")
                || elem_type_param.Contains("узел учета квартиры для хвс")
                || elem_type_param.Contains("узел учета офиса для хвс")
                );
        };

        public CheckHoles() : base()
        {
            if (PluginName == null)
                PluginName = "АР: Проверка отверстий";

            if (ESEntity == null)
                ESEntity = new ExtensibleStorageEntity(
                    PluginName,
                    "KPLN_CheckHoles",
                    new Guid("820080C5-DA99-40D7-9445-E53F288AA160"),
                    new Guid("820080C5-DA99-40D7-9445-E53F288AA161"));
        }

        public override Element[] GetElemsToCheck()
        {
            FamilyInstance[] holesFamInsts;
            if (CheckDocument.Title.Contains("СЕТ_1"))
            {
                holesFamInsts = new FilteredElementCollector(CheckDocument)
                    .OfClass(typeof(FamilyInstance))
                    .OfCategory(BuiltInCategory.OST_Windows)
                    .Cast<FamilyInstance>()
                    .Where(e =>
                        e.Symbol.FamilyName.StartsWith("ASML_АР_Отверстие")
                        && e.GetSubComponentIds().Count == 0)
                    .ToArray();
            }
            else
            {

                holesFamInsts = new FilteredElementCollector(CheckDocument)
                    .OfClass(typeof(FamilyInstance))
                    .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
                    .Cast<FamilyInstance>()
                    .Where(e =>
                        (e.Symbol.FamilyName.StartsWith("199_Отверстие"))
                        && e.GetSubComponentIds().Count == 0)
                    .ToArray();
            }

            return holesFamInsts;
        }

        private protected override CheckResultStatus Set_CheckerEntitiesHeap(Element[] elemColl)
        {
            List<CheckHolesHoleData> holesData = PrepareHoleData(elemColl);
            BoundingBoxXYZ sumBBox = PreparesHolesSumBBox(holesData);
            List<CheckHolesMEPData> mepBBoxData = PrepareMEPData(sumBBox);

            foreach (var hd in holesData)
            {
                hd.SetIntersectsData(mepBBoxData);
            }

            _checkerEntitiesCollHeap = PrepareHolesIntersectsWPFEntity(holesData);
            return CheckResultStatus.Succeeded;
        }

        /// <summary>
        /// Подготовка спец. класса с данными по каждому отверстию
        /// </summary>
        /// <param name="holesColl">Коллекция отверстий</param>
        private List<CheckHolesHoleData> PrepareHoleData(IEnumerable<Element> holesColl)
        {
            List<CheckHolesHoleData> result = new List<CheckHolesHoleData>();

            foreach (Element hole in holesColl)
            {
                CheckHolesHoleData holeData = new CheckHolesHoleData(hole);
                holeData.SetGeometryData(ViewDetailLevel.Coarse);
                if (holeData.CurrentSolid != null)
                    result.Add(holeData);
                else
                    RunElemsWarningColl.Append(new CheckCommandError(hole, $"У элемента с id: {hole.Id} не удалось получить Solid. Проверь отверстие вручную"));
            }

            return result;
        }

        /// <summary>
        /// Создаю основной контур по отверстиям (очень важно для составных моделей)
        /// </summary>
        /// <param name="holesDataColl">Коллекция отверстий CheckHolesHoleData</param>
        private BoundingBoxXYZ PreparesHolesSumBBox(IEnumerable<CheckHolesHoleData> holesDataColl)
        {
            double minX = double.MaxValue;
            double minY = double.MaxValue;
            double minZ = double.MaxValue;
            double maxX = double.MinValue;
            double maxY = double.MinValue;
            double maxZ = double.MinValue;
            foreach (CheckHolesHoleData holeData in holesDataColl)
            {
                #region Получаю минимальную точку в каждой плоскости
                XYZ bboxmim = holeData.CurrentBBox.Min;
                double tminX = Math.Min(minX, bboxmim.X);
                double tminY = Math.Min(minY, bboxmim.Y);
                double tminZ = Math.Min(minZ, bboxmim.Z);
                if (tminX < minX) minX = tminX;
                if (tminY < minY) minY = tminY;
                if (tminZ < minZ) minZ = tminZ;
                #endregion

                #region Получаю максимальную точку в каждой плоскости
                XYZ bboxMax = holeData.CurrentBBox.Max;
                double tmaxX = Math.Max(maxX, bboxMax.X);
                double tmaxY = Math.Max(maxY, bboxMax.Y);
                double tmaxZ = Math.Max(maxZ, bboxMax.Z);
                if (tmaxX > maxX) maxX = tmaxX;
                if (tmaxY > maxY) maxY = tmaxY;
                if (tmaxZ > maxZ) maxZ = tmaxZ;
                #endregion
            }

            return new BoundingBoxXYZ
            {
                Min = new XYZ(minX, minY, minZ),
                Max = new XYZ(maxX, maxY, maxZ)
            };
        }

        /// <summary>
        /// Подготовка коллекции спец. класса с данными по каждому элементу ИОС
        /// </summary>
        /// <param name="bbox">BoundingBoxXYZ, который расчлененного объекта</param>
        private List<CheckHolesMEPData> PrepareMEPData(BoundingBoxXYZ bbox)
        {
            List<CheckHolesMEPData> result = new List<CheckHolesMEPData>();

            IEnumerable<RevitLinkInstance> rvtLinkInsts = new FilteredElementCollector(CheckDocument)
                .OfClass(typeof(RevitLinkInstance))
                // Слабое место - имена файлов могут отличаться из-за требований Заказчика
                .Where(lm =>
                    !(lm.Name.ToUpper().Contains("_KR_")
                        || lm.Name.ToUpper().Contains("_КР_")
                        || (lm.Name.ToUpper().Contains("_KR.rvt") || lm.Name.ToUpper().Contains("_КР.rvt"))
                        || (lm.Name.ToUpper().StartsWith("KR_") || lm.Name.ToUpper().StartsWith("КР_")))
                    && !(lm.Name.ToUpper().Contains("_AR_")
                        || lm.Name.ToUpper().Contains("_АР_"))
                        || (lm.Name.ToUpper().Contains("_AR.RVT") || lm.Name.ToUpper().Contains("_АР.RVT"))
                        || (lm.Name.ToUpper().StartsWith("AR_") || lm.Name.ToUpper().StartsWith("АР_")))
                .Cast<RevitLinkInstance>();

            foreach (RevitLinkInstance rvtLinkInst in rvtLinkInsts)
            {
                Document linkDoc = rvtLinkInst.GetLinkDocument();
                if (linkDoc != null)
                {
                    // Нужна поправка на координаты связей, чтобы сгенерить корретный bbox
                    Transform linkTransform = rvtLinkInst.GetTransform();
                    BoundingBoxXYZ transfBbox = new BoundingBoxXYZ()
                    {
                        Max = linkTransform.Inverse.OfPoint(bbox.Max),
                        Min = linkTransform.Inverse.OfPoint(bbox.Min),
                    };

                    Outline filterOutline = CreateOutlineByBBox(transfBbox);
                    BoundingBoxIntersectsFilter intersectsFilter = new BoundingBoxIntersectsFilter(filterOutline, 0.1);

                    foreach (BuiltInCategory bic in _builtInCategories)
                    {
                        IEnumerable<CheckHolesMEPData> trueMEPElemEntities = new FilteredElementCollector(linkDoc)
                            .OfCategory(bic)
                            .WhereElementIsNotElementType()
                            .WherePasses(intersectsFilter)
                            .Where(_elemExtraFilterFunc)
                            .Cast<Element>()
                            .Select(e => new CheckHolesMEPData(e, rvtLinkInst));

                        List<CheckHolesMEPData> updateMEPElemEntities = new List<CheckHolesMEPData>(trueMEPElemEntities.Count());
                        foreach (CheckHolesMEPData mepElementEntity in trueMEPElemEntities)
                        {
                            #region Блок дополнительной фильтрации
                            if (mepElementEntity.CurrentElement is FamilyInstance mepFI)
                            {
                                // Общий фильтр на вложенные семейства ИОС
                                if (mepFI.SuperComponent != null) continue;

                                // По коннектам - отсеиваю мелкие семейства соединителей, арматуры с подключением < 50 мм. Они 99% попадут по трубе/воздуховоду/лотку. Оборудование - попадает все
                                double tolerance = 0.17;
                                MEPModel mepModel = mepFI.MEPModel;
                                if (mepModel != null
                                    && bic != BuiltInCategory.OST_MechanicalEquipment
                                    && bic != BuiltInCategory.OST_DuctTerminal)
                                {
                                    int isToleranceConCount = 0;
                                    ConnectorManager conManager = mepModel.ConnectorManager;
                                    if (conManager != null)
                                    {
                                        foreach (Connector con in conManager.Connectors)
                                        {
                                            if (con != null && con.Shape == ConnectorProfileType.Round && con.Radius * 2 < tolerance) isToleranceConCount++;
                                            if (con != null && con.Shape == ConnectorProfileType.Rectangular && con.Width < tolerance && con.Height < tolerance) isToleranceConCount++;
                                        }
                                    }
                                    if (isToleranceConCount > 0) continue;
                                }
                            }
                            #endregion

                            #region Блок дополнения элементов геометрией
                            List<XYZ> locationColl = new List<XYZ>(3);
                            Location location = mepElementEntity.CurrentElement.Location;
                            if (location is LocationPoint locationPoint)
                                locationColl.Add(locationPoint.Point);
                            else if (location is LocationCurve locationCurve)
                            {
                                Curve curve = locationCurve.Curve;
                                XYZ start = curve.GetEndPoint(0);
                                locationColl.Add(start);
                                XYZ end = curve.GetEndPoint(1);
                                locationColl.Add(end);
                                XYZ center = new XYZ((start.X + end.X) / 2, (start.Y + end.Y) / 2, (start.Z + end.Z) / 2);
                                locationColl.Add(center);
                            }
                            mepElementEntity.CurrentLocationColl = locationColl;
                            #endregion

                            updateMEPElemEntities.Add(mepElementEntity);
                        }

                        result.AddRange(updateMEPElemEntities);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Создание фильтра, для поиска элементов, с которыми пересекается BoundingBoxXYZ
        /// </summary>
        private Outline CreateOutlineByBBox(BoundingBoxXYZ bbox)
        {
            double minX = bbox.Min.X;
            double minY = bbox.Min.Y;

            double maxX = bbox.Max.X;
            double maxY = bbox.Max.Y;

            double sminX = Math.Min(minX, maxX);
            double sminY = Math.Min(minY, maxY);

            double smaxX = Math.Max(minX, maxX);
            double smaxY = Math.Max(minY, maxY);

            XYZ pntMax = new XYZ(smaxX, smaxY, bbox.Max.Z);
            XYZ pntMin = new XYZ(sminX, sminY, bbox.Min.Z);

            return new Outline(pntMin, pntMax);
        }

        /// <summary>
        /// Подготовка WPFEntity по отверстий, в зависимости от пересечений с элементами ИОС
        /// </summary>
        /// <param name="holesData">Коллеция спец. классов</param>
        private List<CheckerEntity> PrepareHolesIntersectsWPFEntity(List<CheckHolesHoleData> holesData)
        {
            List<CheckerEntity> result = new List<CheckerEntity>();

            foreach (CheckHolesHoleData holeData in holesData)
            {
                Element hole = holeData.CurrentElement;
                Level holeLevel = (Level)CheckDocument.GetElement(hole.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM).AsElementId());

                if (holeData.IntesectElementsColl.Count() == 0)
                {
                    CheckerEntity zeroIOSElem = new CheckerEntity(
                        hole,
                        "Отверстие не содержит элементов ИОС",
                        $"Отверстие должно быть заполнено элементами ИОС, иначе оно лишнее",
                        $"Ошибка может быть ложной, если не все связи ИОС загружены в проект, " +
                        $"ЛИБО отверстие выдано инженерами по отдельному заданию (например - лючок под коллектор ОВ).\nУровень размещения: {holeLevel.Name}",
                        true)
                        .Set_CanApprovedAndESData(ESEntity)
                        .Set_ZoomData(holeData.CurrentBBox);

                    result.Add(zeroIOSElem);
                    continue;
                }

                double intersectPersent = holeData.SumIntersectArea / holeData.MainHoleFace.Area;
                if (intersectPersent < 0.200 && !holeData.IntesectElementsColl.Any(hd => hd.CurrentElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeCurves))
                {
                    CheckerEntity errorNoPipeAreaElem = new CheckerEntity(
                        hole,
                        "Отверстие избыточное по размерам",
                        $"Большая вероятность, что необходимо пересмотреть размеры, т.к. отверстие без труб, и заполнено элементами ИОС только на {Math.Round(intersectPersent, 3) * 100}%.",
                        $"Ошибка может быть ложной, если не все связи ИОС загружены в проект.\nУровень размещения: {holeLevel.Name}",
                        true)
                        .Set_CanApprovedAndESData(ESEntity, ErrorStatus.Warning)
                        .Set_ZoomData(holeData.CurrentBBox);

                    result.Add(errorNoPipeAreaElem);
                    continue;
                }
                else if (intersectPersent < 0.250 && holeData.IntesectElementsColl.Count() == 1)
                {
                    CheckerEntity errorOneElemAreaElem = new CheckerEntity(
                        hole,
                        "Отверстие избыточное по размерам",
                        $"Большая вероятность, что необходимо пересмотреть размеры, т.к. отверстие заполнено 1 элементом ИОС на {Math.Round(intersectPersent, 3) * 100}%.",
                        $"Ошибка может быть ложной, если не все связи ИОС загружены в проект.\nУровень размещения: {holeLevel.Name}",
                        true)
                        .Set_CanApprovedAndESData(ESEntity, ErrorStatus.Warning)
                        .Set_ZoomData(holeData.CurrentBBox);

                    result.Add(errorOneElemAreaElem);
                    continue;
                }
                else if (intersectPersent < 0.150)
                {
                    CheckerEntity warnAreaElem = new CheckerEntity(
                        hole,
                        "Отверстие избыточное по размерам",
                        $"Возможно стоит пересмотреть размеры, т.к. отверстие заполнено элементами ИОС только на {Math.Round(intersectPersent, 3) * 100}%.",
                        $"Ошибка может быть ложной, если не все связи ИОС загружены в проект.\nУровень размещения: {holeLevel.Name}",
                        true)
                        .Set_CanApprovedAndESData(ESEntity, ErrorStatus.Warning)
                        .Set_ZoomData(holeData.CurrentBBox);

                    result.Add(warnAreaElem);
                    continue;
                }
            }

            return result
                .OrderBy(e =>
                    ((Level)CheckDocument.GetElement(e.Element.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM).AsElementId())).Elevation)
                .ToList();
        }
    }
}
