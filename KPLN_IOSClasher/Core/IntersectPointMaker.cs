using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;

namespace KPLN_IOSClasher.Core
{
    /// <summary>
    /// Сущность для работы с точками пересечений
    /// </summary>
    internal class IntersectPointMaker
    {
        public static readonly string ClashPointFamilyName = "ClashPoint_Small";

        public static readonly Guid PointCoord_Param = new Guid("76b464e2-b02e-4af8-bd1a-eb4985a0f362");
        public static readonly Guid AddedElementId_Param = new Guid("bf5ddf52-e83e-4094-9baf-db5686356b5d");
        public static readonly Guid OldElementId_Param = new Guid("dc3f91c3-8144-4e2c-9b4f-ea060706a2ea");
        public static readonly Guid LinkInstanceId_Param = new Guid("468b7905-8f4e-4f72-86f3-3a18685a3838");
        public static readonly Guid UserData_Param = new Guid("41b103e9-3eb8-43fe-af9d-b74a13acf45b");
        public static readonly Guid CurrentData_Param = new Guid("8984f2e6-606d-4cd0-8fb7-19bb034c5059");
        public static readonly Guid CurrentDiameter_Param = new Guid("9b679ab7-ea2e-49ce-90ab-0549d5aa36ff");

        private static readonly Guid _famVersion_Param = new Guid("85cd0032-c9ee-4cd3-8ffa-b2f1a05328e3");
        private static readonly string _famVersion_Data = "1.1";
        private static FamilySymbol _clashPointFamilySymbol;

        protected static XYZ ParseStringToXYZ(string input)
        {
            try
            {
                // Выдаляем дужкі і разбіваем радок па косцы
                string cleanedInput = input.Trim('(', ')');
                string[] parts = cleanedInput.Split(',');

                // Парсим кожную частку ў double
                double x = double.Parse(parts[0], CultureInfo.InvariantCulture);
                double y = double.Parse(parts[1], CultureInfo.InvariantCulture);
                double z = double.Parse(parts[2], CultureInfo.InvariantCulture);

                // Ствараем і вяртаем XYZ
                return new XYZ(x, y, z);
            }
            catch (Exception ex)
            {
                // Абработка памылак парсінгу
                throw new FormatException($"Отправь разработчику: Не удалось конвертировать данные из строки в XYZ", ex);
            }
        }

        /// <summary>
        /// Получить коллекцию уже РАМЗЕЩЕННЫХ элементов пересечений
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        protected static Element[] GetOldIntersectionPointsEntities(Document doc) =>
            new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .Where(el => el is FamilyInstance famInst && famInst.Symbol.FamilyName == ClashPointFamilyName)
                .ToArray();

        /// <summary>
        /// Очистка коллекции на создание от близких соседей
        /// </summary>
        protected static IntersectPointEntity[] ClearedNeighbourEntities(IEnumerable<Element> oldPointElems, IEnumerable<IntersectPointEntity> intersectPointEntities)
        {
            Element[] onlyValidElems = oldPointElems.Where(el => el.IsValidObject).ToArray();

            if (onlyValidElems.Length == 0)
                return intersectPointEntities.ToArray();

            IntersectPointEntity[] clearedElem = intersectPointEntities
                .Where(ipe =>
                    onlyValidElems.All(ve => Math.Abs(ipe.IntersectPoint.DistanceTo(ParseStringToXYZ(ve.get_Parameter(PointCoord_Param).AsString()))) > 0.05))
                .ToArray();

            return clearedElem;
        }

        /// <summary>
        /// Разместить экземпляр семейства пересечения по указанным координатам
        /// </summary>
        protected static void CreateIntersectFamilyInstance(Document doc, IntersectPointEntity[] clearedPointEntities)
        {
            FamilySymbol intersectFamSymb = GetIntersectFamilySymbol(doc);

            // Создание новых по уточненной коллекции
            foreach (IntersectPointEntity entity in clearedPointEntities)
            {
                XYZ point = entity.IntersectPoint;
                Level level = GetNearestLevel(doc, point.Z) ?? throw new Exception("В проекте отсутсвуют уровни!");

                FamilyInstance instance = doc
                    .Create
                    .NewFamilyInstance(point, intersectFamSymb, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                doc.Regenerate();

                // Указать уровень
                instance.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(point.Z - level.Elevation);

                // Указать данные по коллизии
                instance.get_Parameter(AddedElementId_Param).Set(entity.AddedElement_Id.ToString());
                instance.get_Parameter(OldElementId_Param).Set(entity.OldElement_Id.ToString());
                instance.get_Parameter(LinkInstanceId_Param).Set(entity.LinkInstance_Id.ToString());
                instance.get_Parameter(PointCoord_Param).Set(point.ToString());
                instance.get_Parameter(UserData_Param).Set($"{entity.CurrentUser.Name} {entity.CurrentUser.Surname}");
                instance.get_Parameter(CurrentData_Param).Set(DateTime.Now.ToString("g"));
                instance.get_Parameter(CurrentDiameter_Param).Set(CreateIntersectDiameter(entity.IntersectSolid));

                doc.Regenerate();
            }
        }

        /// <summary>
        /// Получить семейство отображающее коллизию
        /// </summary>
        private static FamilySymbol GetIntersectFamilySymbol(Document doc)
        {
#if Debug
            _clashPointFamilySymbol = null;
#endif

            if (_clashPointFamilySymbol == null)
            {
                FamilySymbol[] oldFamSymbOfGM = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_GenericModel)
                    .Where(el => el is FamilySymbol famSymb && famSymb.FamilyName == ClashPointFamilyName)
                    .Cast<FamilySymbol>()
                    .ToArray();

                bool versionVerifyError = false;
                if (oldFamSymbOfGM.Any())
                    versionVerifyError = FamilySymboVersionErrorCheck(oldFamSymbOfGM.FirstOrDefault());

                // Если в проекте нет, или была ошибка версии - то грузим\обновляем
                if (!oldFamSymbOfGM.Any() || versionVerifyError)
                {
                    string path = $@"{Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)}\Source\RevitData\{Module.RevitVersion}\{ClashPointFamilyName}.rfa";
                    bool result = doc.LoadFamily(path, new FamilyLoadOptions(), out Family fam);
                    if (!result)
                        throw new Exception("Семейство для метки не найдено! Обратись к разработчику.");

                    doc.Regenerate();

                    _clashPointFamilySymbol = doc.GetElement(fam.GetFamilySymbolIds().FirstOrDefault()) as FamilySymbol;
                }
                else
                    _clashPointFamilySymbol = oldFamSymbOfGM.FirstOrDefault();
            }

            _clashPointFamilySymbol.Activate();

            return _clashPointFamilySymbol;
        }

        private static bool FamilySymboVersionErrorCheck(FamilySymbol checkFS)
        {
            Parameter famVersion = checkFS.get_Parameter(_famVersion_Param);
            if (famVersion == null)
                return true;

            string famVersionData = famVersion.AsString();
            if (!string.IsNullOrEmpty(famVersionData) && famVersionData != _famVersion_Data)
                return true;

            return false;
        }

        /// <summary>
        /// Создание значение диаметра для клэшпоинта
        /// </summary>
        /// <param name="solid"></param>
        /// <returns></returns>
        private static double CreateIntersectDiameter(Solid solid)
        {
            double resultDiam = 2;

            if(solid != null)
            {
                double solidVolume = solid.Volume;
                if (solidVolume <= 1)
                    resultDiam = 0.5;
                else if (solidVolume > 1 && solidVolume <= 3)
                    resultDiam = 1;
                else if (solidVolume > 3 && solidVolume <= 10)
                    resultDiam = 2;
                else if (solidVolume > 10 && solidVolume <= 20)
                    resultDiam = 3;
                else if (solidVolume > 20 && solidVolume <= 30)
                    resultDiam = 4;
                else
                    resultDiam = 6;
            }

            return resultDiam;
        }

        /// <summary>
        ///  Поиск ближайшего подходящего уровня
        /// </summary>
        private static Level GetNearestLevel(Document doc, double elevation)
        {
            Level result = null;

            double resultDistance = 999999;
            foreach (Level lvl in GetLevels(doc))
            {
                double tempDistance = Math.Abs(lvl.Elevation - elevation);
                if (Math.Abs(lvl.Elevation - elevation) < resultDistance)
                {
                    result = lvl;
                    resultDistance = tempDistance;
                }
            }
            return result;
        }

        /// <summary>
        /// Получить коллекцию ВСЕХ уровней проекта
        /// </summary>
        private static Level[] GetLevels(Document doc)
        {
            Level[] instances = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Levels)
                .WhereElementIsNotElementType()
                .Cast<Level>()
                .ToArray();

            return instances;
        }
    }
}
