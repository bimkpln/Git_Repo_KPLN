using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_PluginActivityWorker;
using KPLN_ModelChecker_Lib;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_ModelChecker_User.Common.CheckCommandCollections;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandCheckHoles : AbstrCheckCommand<CommandCheckHoles>, IExternalCommand
    {
        internal const string PluginName = "АР: Проверка отверстий";

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

        public CommandCheckHoles() : base()
        {
        }

        internal CommandCheckHoles(ExtensibleStorageEntity esEntity) : base(esEntity)
        {
        }

        /// <summary>
        /// Реализация IExternalCommand
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return ExecuteByUIApp(commandData.Application);
        }

        public override Result ExecuteByUIApp(UIApplication uiapp)
        {
            DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName($"{PluginName}", ModuleData.ModuleName).ConfigureAwait(false);

            _uiApp = uiapp;

            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Получаю коллекцию элементов для анализа
            FamilyInstance[] holesFamInsts = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
                .Cast<FamilyInstance>()
                .Where(e =>
                    e.Symbol.FamilyName.StartsWith("199_Отверстие")
                    && e.GetSubComponentIds().Count == 0)
                .ToArray();

            #region Проверяю и обрабатываю элементы
            WPFEntity[] wpfColl = CheckCommandRunner(doc, holesFamInsts);
            OutputMainForm form = ReportCreatorAndDemonstrator(doc, wpfColl);
            if (form != null) form.Show();
            else return Result.Cancelled;
            #endregion

            return Result.Succeeded;
        }

        private protected override IEnumerable<CheckCommandError> CheckElements(Document doc, object[] objColl)
        {
            if (!(objColl.Any()))
                throw new CheckerException("Не удалось определить семейства. Поиск осуществялется по категории 'Оборудование', и имени, которое начинается с '199_Отверстие'");

            return Enumerable.Empty<CheckCommandError>();
        }

        private protected override IEnumerable<WPFEntity> PreapareElements(Document doc, Element[] elemColl)
        {
            List<CheckHolesHoleData> holesData = PrepareHoleData(elemColl);
            BoundingBoxXYZ sumBBox = PreparesHolesSumBBox(holesData);
            List<CheckHolesMEPData> mepBBoxData = PrepareMEPData(doc, sumBBox);

            foreach (var hd in holesData)
            {
                hd.SetIntersectsData(mepBBoxData);
            }

            return PrepareHolesIntersectsWPFEntity(doc, holesData);
        }

        private protected override void SetWPFEntityFiltration(WPFReportCreator report)
        {
            report.SetWPFEntityFiltration_ByStatus();
        }

        /// <summary>
        /// Подготовка спец. класса с данными по каждому отверстию
        /// </summary>
        /// <param name="holesFamInsts">Коллекция отверстий</param>
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
                    _errorRunColl.Append(new CheckCommandError(hole, $"У элемента с id: {hole.Id} не удалось получить Solid. Проверь отверстие вручную"));
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
        /// <param name="doc">Документ rvt</param>
        /// <param name="bbox">BoundingBoxXYZ, который расчлененного объекта</param>
        private List<CheckHolesMEPData> PrepareMEPData(Document doc, BoundingBoxXYZ bbox)
        {
            List<CheckHolesMEPData> result = new List<CheckHolesMEPData>();

            IEnumerable<RevitLinkInstance> rvtLinkInsts = new FilteredElementCollector(doc)
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

                    BoundingBoxIntersectsFilter filter = CreateFilter(transfBbox);
                    foreach (BuiltInCategory bic in _builtInCategories)
                    {
                        IEnumerable<CheckHolesMEPData> trueMEPElemEntities = new FilteredElementCollector(linkDoc)
                            .OfCategory(bic)
                            .WhereElementIsNotElementType()
                            .WherePasses(filter)
                            .Cast<Element>()
                            .Select(e => new CheckHolesMEPData(e, rvtLinkInst));

                        List<CheckHolesMEPData> updateMEPElemEntities = new List<CheckHolesMEPData>(trueMEPElemEntities.Count());
                        foreach (CheckHolesMEPData mepElementEntity in trueMEPElemEntities)
                        {
                            #region Блок дополнительной фильтрации
                            if (mepElementEntity.CurrentElement is FamilyInstance mepFI)
                            {
                                // Общий фильтр на вложенные семейства, и на семейства отверстий ЗИ ИОС
                                if (mepFI.Symbol.FamilyName.StartsWith("501_") || mepFI.SuperComponent != null) continue;

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
        /// Подготовка WPFEntity по отверстий, в зависимости от пересечений с элементами ИОС
        /// </summary>
        /// <param name="holesData">Коллеция спец. классов</param>
        private List<WPFEntity> PrepareHolesIntersectsWPFEntity(Document doc, List<CheckHolesHoleData> holesData)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            foreach (CheckHolesHoleData holeData in holesData)
            {
                Element hole = holeData.CurrentElement;
                Level holeLevel = (Level)doc.GetElement(hole.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM).AsElementId());

                if (holeData.IntesectElementsColl.Count() == 0)
                {
                    WPFEntity zeroIOSElem = new WPFEntity(
                        ESEntity,
                        hole,
                        "Отверстие не содержит элементов ИОС",
                        $"Отверстие должно быть заполнено элементами ИОС, иначе оно лишнее",
                        $"Ошибка может быть ложной, если не все связи ИОС загружены в проект.\nУровень размещения: {holeLevel.Name}",
                        true,
                        true);

                    zeroIOSElem.ResetZoomGeometryExtension(holeData.CurrentBBox);
                    result.Add(zeroIOSElem);
                    continue;
                }

                double intersectPersent = holeData.SumIntersectArea / holeData.MainHoleFace.Area;
                if (intersectPersent < 0.200 && !holeData.IntesectElementsColl.Any(hd => hd.CurrentElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeCurves))
                {
                    WPFEntity errorNoPipeAreaElem = new WPFEntity(
                        ESEntity,
                        hole,
                        "Отверстие избыточное по размерам",
                        $"Большая вероятность, что необходимо пересмотреть размеры, т.к. отверстие без труб, и заполнено элементами ИОС только на {Math.Round(intersectPersent, 3) * 100}%.",
                        $"Ошибка может быть ложной, если не все связи ИОС загружены в проект.\nУровень размещения: {holeLevel.Name}",
                        true,
                        CheckStatus.Warning,
                        true);

                    errorNoPipeAreaElem.ResetZoomGeometryExtension(holeData.CurrentBBox);
                    result.Add(errorNoPipeAreaElem);
                    continue;
                }
                else if (intersectPersent < 0.250 && holeData.IntesectElementsColl.Count() == 1)
                {
                    WPFEntity errorOneElemAreaElem = new WPFEntity(
                        ESEntity,
                        hole,
                        "Отверстие избыточное по размерам",
                        $"Большая вероятность, что необходимо пересмотреть размеры, т.к. отверстие заполнено 1 элементом ИОС на {Math.Round(intersectPersent, 3) * 100}%.",
                        $"Ошибка может быть ложной, если не все связи ИОС загружены в проект.\nУровень размещения: {holeLevel.Name}",
                        true,
                        CheckStatus.Warning,
                        true);

                    errorOneElemAreaElem.ResetZoomGeometryExtension(holeData.CurrentBBox);
                    result.Add(errorOneElemAreaElem);
                    continue;
                }
                else if (intersectPersent < 0.150)
                {
                    WPFEntity warnAreaElem = new WPFEntity(
                        ESEntity,
                        hole,
                        "Отверстие избыточное по размерам",
                        $"Возможно стоит пересмотреть размеры, т.к. отверстие заполнено элементами ИОС только на {Math.Round(intersectPersent, 3) * 100}%.",
                        $"Ошибка может быть ложной, если не все связи ИОС загружены в проект.\nУровень размещения: {holeLevel.Name}",
                        true,
                        CheckStatus.Warning,
                        true);

                    warnAreaElem.ResetZoomGeometryExtension(holeData.CurrentBBox);
                    result.Add(warnAreaElem);
                    continue;
                }
            }

            return result
                .OrderBy(e =>
                    ((Level)doc.GetElement(e.Element.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM).AsElementId())).Elevation)
                .ToList();
        }

        /// <summary>
        /// Создание фильтра, для поиска элементов, с которыми пересекается BoundingBoxXYZ
        /// </summary>
        private BoundingBoxIntersectsFilter CreateFilter(BoundingBoxXYZ bbox)
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

            Outline outline = new Outline(pntMin, pntMax);

            return new BoundingBoxIntersectsFilter(outline);
        }
    }
}
