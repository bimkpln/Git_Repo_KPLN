using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_Lib.Services.GripGeom.Core
{
    /// <summary>
    /// Класс для генерации солида между уровнем и ограждающими секциями
    /// </summary>
    public sealed class LevelAndSectionSolid
    {
        private static readonly Guid _lvlParamGuid = new Guid("9eabf56c-a6cd-4b5c-a9d0-e9223e19ea3f");
        private static readonly Guid _sectParamGuid = new Guid("b3aaab47-69d0-4e1c-9d70-c0d6907961cc");
        private static readonly Guid _corpsParamGuid = new Guid("3f509328-1c21-4bf8-a156-dc71ec7242fd");

        private static readonly Guid _upLvlParam = new Guid("a5a5776e-f3db-4755-bb23-88ccb9632ced");
        private static readonly Guid _downLvlParam = new Guid("c033134a-9b0c-49c9-b389-8a94314ddedb");

        /// <summary>
        /// Элемент из Revit в границах уровней и секций
        /// </summary>
        public Element LSElement { get; private set; }

        /// <summary>
        /// Solid в границах уровней и секций
        /// </summary>
        public Solid LSSolid { get; private set; }

        /// <summary>
        /// Значение параметра уровня
        /// </summary>
        public string LSLevelData { get; private set; }

        /// <summary>
        /// Значение параметра секции
        /// </summary>
        public string LSSectionData { get; private set; }

        /// <summary>
        /// Значение параметра корпуса
        /// </summary>
        public string LSCorpsData { get; private set; }

        /// <summary>
        /// Физическая привязка к уровню выше. Null - если бесконечность (уровень кровли)
        /// </summary>
        public Level LSUpLevelFromModel { get; private set; }

        /// <summary>
        /// Физическая привязка к уровню ниже. Null - если бесконечность (нижний уровень)
        /// </summary>
        public Level LSDownLevelFromModel { get; private set; }

        /// <summary>
        /// BBox в границах уровней и секций
        /// </summary>
        public BoundingBoxXYZ LSBBox { get; private set; }

        /// <summary>
        /// BBox в границах уровней и секций
        /// </summary>
        public Outline BBoxOutline { get; private set; }

        private LevelAndSectionSolid(Element elem, Solid solid, Level upLevelFromModel, Level downLevelFromModel, string levelData, string sectionData)
        {
            LSElement = elem;
            LSSolid = solid;
            LSSectionData = sectionData;
            LSLevelData = levelData;
            LSUpLevelFromModel = upLevelFromModel;
            LSDownLevelFromModel = downLevelFromModel;

            LSBBox = GeometryWorker.GetBoundingBoxXYZ(LSSolid);
            BBoxOutline = GeometryWorker.CreateOutline_ByBBoxANDExpand(LSBBox, new XYZ(2, 2, 2));
        }

        private LevelAndSectionSolid(Element elem, Solid solid, Level upLevelFromModel, Level downLevelFromModel, string levelData, string sectionData, string corpsData) : this (elem, solid, upLevelFromModel, downLevelFromModel, levelData, sectionData)
        {
            LSCorpsData = corpsData;
        }

        /// <summary>
        /// Подготовка коллекции солидов с данными по секциям
        /// </summary>
        /// <param name="doc">Revit-документ</param>
        public static List<LevelAndSectionSolid> PrepareSolids(Document doc, bool isCorpsData)
        {
            List<LevelAndSectionSolid> result = new List<LevelAndSectionSolid>();


            // Собираю линки
            RevitLinkInstance[] rlikColl = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .WhereElementIsNotElementType()
                .Cast<RevitLinkInstance>()
                .Where(rli =>
                    (rli.Name.Contains("Разб.Файл") || rli.Name.Contains("Разбив") || rli.Name.Contains("РазбФайл") || rli.Name.Contains("РФ"))
                    && rli.GetLinkDocument() != null)
                .ToArray();

            if (rlikColl.Length == 0)
                throw new CheckerException($"Прервано по причине отсутствия разбивочного файла в модели.\n" +
                    $"Разбивочную модуль нужно добавить связью и обязательно открыть");

            // Остортированная коллекция уровней
            List<Level> orderedLvls = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Levels)
                .WhereElementIsNotElementType()
                .Cast<Level>()
                .OrderBy(lvl => lvl.Elevation)
                .ToList();


            // Анализ линков и генерация результата
            ElementClassFilter classFilter = new ElementClassFilter(typeof(FamilyInstance));
            foreach (RevitLinkInstance rli in rlikColl)
            {
                Document linkDoc = rli.GetLinkDocument();
                if (linkDoc == null)
                    continue;


                FamilyInstance[] lsLinkMainFIColl = new FilteredElementCollector(linkDoc)
                    .OfCategory(BuiltInCategory.OST_Mass)
                    .OfClass(typeof(FamilyInstance))
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol.FamilyName.Contains("001_Разделитель захваток"))
                    .ToArray();


                // Беру трансформ для линка
                Transform lDocTrans = rli.GetTotalTransform();
                //if (!lDocTrans.IsTranslation)
                //    lDocTrans = Transform.CreateTranslation(lDocTrans.Origin);

                foreach (FamilyInstance lsMainFI in lsLinkMainFIColl)
                {
                    Element[] subElems = lsMainFI
                        .GetSubComponentIds()
                        .Select(id => linkDoc.GetElement(id))
                        .ToArray();

                    foreach (Element elem in subElems)
                    {
                        Parameter cElemParam = elem.get_Parameter(_corpsParamGuid);
                        Parameter sElemParam = elem.get_Parameter(_sectParamGuid);
                        Parameter lElemParam = elem.get_Parameter(_lvlParamGuid);
                        Parameter upLvlParam = elem.get_Parameter(_upLvlParam);
                        Parameter downLvlParam = elem.get_Parameter(_downLvlParam);

                        if (isCorpsData && (cElemParam == null || sElemParam == null || lElemParam == null || upLvlParam == null || downLvlParam == null))
                            throw new CheckerException($"У элемента с id: {elem.Id} из модели {linkDoc.Title} нет одного из параметров. " +
                                $"Проверь наличие: 1. КП_О_Этаж; 2. КП_О_Секция;  3. КП_О_Корпус; 4. ПЗ_Верхний уровень; 5. ПЗ_Нижний уровень");
                        else if (sElemParam == null || lElemParam == null || upLvlParam == null || downLvlParam == null)
                            throw new CheckerException($"У элемента с id: {elem.Id} из модели {linkDoc.Title} нет одного из параметров. " +
                                $"Проверь наличие: 1. КП_О_Этаж; 2. КП_О_Секция; 3. ПЗ_Верхний уровень; 4. ПЗ_Нижний уровень");

                        Solid elemSolid = GeometryWorker.GetRevitElemUniontSolid(elem, lDocTrans)
                            ?? throw new CheckerException($"У элемента с id: {elem.Id} не удалось определить Solid. Отправь разработчику");


                        Level upLvl = null;
                        string upLvlParamData = upLvlParam.AsString() 
                            ?? throw new CheckerException($"У элемента с id: {elem.Id} не заполнен параметр уровня \"ПЗ_Верхний уровень\". Отправь координатору на доработку");
                        if (double.TryParse(upLvlParamData, out double upLvlParamDataDouble))
                            upLvl = LevelWorker.BinaryFindExactLevel(orderedLvls, upLvlParamDataDouble);
                        
                        Level downLvl = null;
                        string downLvlParamData = downLvlParam.AsString()
                            ?? throw new CheckerException($"У элемента с id: {elem.Id} не заполнен параметр уровня \"ПЗ_Нижний уровень\". Отправь координатору на доработку");
                        if (double.TryParse(downLvlParamData, out double downLvlParamDataDouble))
                            downLvl = LevelWorker.BinaryFindExactLevel(orderedLvls, downLvlParamDataDouble);

                        // Допускаю null в значениях уровней. Такое может быть для граничных боксов
                        if (isCorpsData)
                            result.Add(new LevelAndSectionSolid(elem, elemSolid, upLvl, downLvl, lElemParam.AsString(), sElemParam.AsString(), cElemParam.AsString()));
                        else
                            result.Add(new LevelAndSectionSolid(elem, elemSolid, upLvl, downLvl, lElemParam.AsString(), sElemParam.AsString()));
                    }
                }
            }


            if (result.Count() == 0)
                throw new CheckerException($"Прервано по причине отсутствия в разбивочном файле экземпляра семейства для разделения по захваткам.\n" +
                    $"Напиши своему координатору");

            return result;
        }
    }
}
