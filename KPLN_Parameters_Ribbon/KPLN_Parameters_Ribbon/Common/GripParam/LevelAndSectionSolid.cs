using Autodesk.Revit.DB;
using KPLN_Parameters_Ribbon.Services;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Parameters_Ribbon.Common.GripParam
{
    /// <summary>
    /// Класс для генерации солида между уровнем и ограждающими секциями
    /// </summary>
    public class LevelAndSectionSolid
    {
        private static readonly Guid _lvlParamGuid = new Guid("9eabf56c-a6cd-4b5c-a9d0-e9223e19ea3f");
        private static readonly Guid _sectParamGuid = new Guid("b3aaab47-69d0-4e1c-9d70-c0d6907961cc");
        
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
        /// BBox в границах уровней и секций
        /// </summary>
        public BoundingBoxXYZ LSBBox { get; private set; }

        /// <summary>
        /// BBox в границах уровней и секций
        /// </summary>
        public Outline BBoxOutline { get; private set; }

        private LevelAndSectionSolid(Element elem, Solid solid, string sectionData, string levelData)
        {
            LSElement = elem;
            LSSolid = solid;
            LSSectionData = sectionData;
            LSLevelData = levelData;

            LSBBox = GeometryWorker.GetBoundingBoxXYZ(LSSolid);
            BBoxOutline = GeometryWorker.CreateOutline_ByBBoxANDExpand(LSBBox, 2);
        }

        /// <summary>
        /// Подготовка коллекции солидов с данными по секциям
        /// </summary>
        /// <param name="doc">Revit-документ</param>
        /// <param name="sectSeparParamName">Параметр для по секциям</param>
        /// <param name="lvlSeparParamName">Параметр для по уровням</param>
        public static List<LevelAndSectionSolid> PrepareSolids(Document doc, string sectSeparParamName, string lvlSeparParamName)
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
                throw new GripParamExection($"Прервано по причине отсутствия разбивочного файла в модели.\n" +
                    $"Разбивочную модуль нужно добавить связью и обязательно открыть");


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
                Transform lDocTrans = rli.GetTransform();
                if (!lDocTrans.IsTranslation)
                    lDocTrans = Transform.CreateTranslation(lDocTrans.Origin);

                foreach (FamilyInstance lsMainFI in lsLinkMainFIColl)
                {
                    Element[] subElems = lsMainFI
                        .GetSubComponentIds()
                        .Select(id => linkDoc.GetElement(id))
                        .ToArray();

                    foreach (Element elem in subElems)
                    {
                        Parameter sElemParam = elem.get_Parameter(_sectParamGuid);
                        Parameter lElemParam = elem.get_Parameter(_lvlParamGuid);
                        if (sElemParam == null || lElemParam == null)
                            continue;

                        Solid elemSolid = GeometryWorker.GetRevitElemUniontSolid(elem, lDocTrans);
                        if (elemSolid == null)
                            continue;

                        result.Add(new LevelAndSectionSolid(elem, elemSolid, sElemParam.AsString(), lElemParam.AsString()));
                    }
                }
            }


            if (result.Count() == 0)
                throw new GripParamExection($"Прервано по причине отсутствия в разбивочном файле экземпляра семейства для разделения по захваткам.\n" +
                    $"Напиши своему координатору");

            return result;
        }
    }
}

