using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_IOSClasher.Core;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;

namespace KPLN_IOSClasher.ExecutableCommand
{
    /// <summary>
    /// Класс для СОЗДАНИЯ и ПЕРЕСТРОЕНИЯ элементов пересечений
    /// </summary>
    internal class IntersectPointMaker : IntersectPointFamInst, IExecutableCommand
    {
        /// <summary>
        /// Имя транзакции для анализа на наличие
        /// </summary>
        public static readonly string TransName = "KPLN: Создание меток коллизий";

        private static string _revitVersion;
        private readonly IntersectPointEntity[] _intersectPointEntities;
        private readonly Element[] _checkedElems;

        public IntersectPointMaker(IEnumerable<IntersectPointEntity> intersectPoints, IEnumerable<Element> checkedElems)
        {
            CultureInfo.CurrentCulture = new CultureInfo("ru-RU");

            _intersectPointEntities = intersectPoints.ToArray();
            _checkedElems = checkedElems.ToArray();
        }

        public Result Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;
            UIDocument uidoc = app.ActiveUIDocument;
            if (string.IsNullOrEmpty(_revitVersion))
                _revitVersion = app.Application.VersionNumber;

            if (uidoc == null)
                return Result.Cancelled;

            using (Transaction trans = new Transaction(doc, TransName))
            {
                trans.Start();

                Element[] oldPointElems = GetOldIntersectionPointsEntities(doc);
                IntersectPointEntity[] clearedPointEntities = null;
                // Удаляю не актуальные
                if (oldPointElems.Length != 0)
                {
                    // Удаляю не актуальные (если можно их удалить). Проблема с занятами клэшпоинтами приводит к ложным клэшам. В пределах погрешности ок, ведь когда 
                    // юзер зайдет в модель - он свои клэши почистит (для него эти эл-ты уже не заняты)
                    List<Element> oldElemsToDel = GetOldToDelete(doc, oldPointElems, _checkedElems);
                    ICollection<ElementId> availableWSElemsId = WorksharingUtils.CheckoutElements(doc, oldElemsToDel.Select(el => el.Id).ToArray());
                    doc.Delete(availableWSElemsId);

                    clearedPointEntities = ClearedNewEntities(oldPointElems, _intersectPointEntities);
                }
                else
                    clearedPointEntities = _intersectPointEntities;

                // Создаю новые
                CreateIntersectFamilyInstance(doc, clearedPointEntities);

                trans.Commit();

                if (_intersectPointEntities.Any())
                {
                    if (!KPLN_Loader.Application.OnIdling_CommandQueue.Any(i => i is IntersectPointMaker))
                    {
                        TaskDialog td = new TaskDialog("ВНИМАНИЕ: Коллизии")
                        {
                            MainIcon = TaskDialogIcon.TaskDialogIconError,
                            MainInstruction = "При создании/редактировании линейной сети вы создали коллизии с другими сетями. ",
                            CommonButtons = TaskDialogCommonButtons.Ok,
                        };

                        td?.Show();
                    }
                }
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Получить коллекцию уже РАМЗЕЩЕННЫХ элементов пересечений
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        public static Element[] GetOldIntersectionPointsEntities(Document doc) =>
            new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .Where(el => el is FamilyInstance famInst && famInst.Symbol.FamilyName == ClashPointFamilyName)
                .ToArray();

        /// <summary>
        /// Очистка коллекции на создание от уже существующих точек (по координатам)
        /// </summary>
        public static IntersectPointEntity[] ClearedNewEntities(IEnumerable<Element> oldPointElems, IEnumerable<IntersectPointEntity> intersectPointEntities)
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
        /// Подготовка списка на удаление для элементов, которые были удалены (по Id)
        /// </summary>
        public static List<Element> GetOldToDelete(Document doc, IEnumerable<Element> oldPointElems, IEnumerable<Element> checkedElems)
        {
            List<Element> resultToDel = new List<Element>();

            foreach (Element oldElem in oldPointElems)
            {
                int oldAddedElemId = -1;
                int oldOldElemId = -1;
                try
                {
                    oldAddedElemId = int.Parse(oldElem.get_Parameter(AddedElementId_Param).AsString());
                    oldOldElemId = int.Parse(oldElem.get_Parameter(OldElementId_Param).AsString());
                }
                catch (Exception ex)
                {
                    throw new FormatException($"Отправь разработчику: Не удалось конвертировать данные из id-элементов для эл-та с id:{oldElem.Id}, причина {ex.Message}");
                }

                if (checkedElems.Any(e => e.Id.IntegerValue == oldAddedElemId || e.Id.IntegerValue == oldOldElemId))
                    resultToDel.Add(oldElem);

                // Очистка от ложных элементов, которые могли остаться после оперирования командами Redo/Undo, и как следствие - клэшпоинт остался, а элемента может не быть. 
                // Блок с очисткой клэшпоинта по существующему элементу но с другим координатами - смотри в ClearedNewEntities блок с проверкой координат
                if (doc.GetElement(new ElementId(oldAddedElemId)) == null)
                    resultToDel.Add(oldElem);
            }

            return resultToDel;
        }

        /// <summary>
        /// Разместить экземпляр семейства пересечения по указанным координатам
        /// </summary>
        public static void CreateIntersectFamilyInstance(Document doc, IntersectPointEntity[] clearedPointEntities)
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

                doc.Regenerate();
            }
        }

        /// <summary>
        /// Получить семейство отображающее коллизию
        /// </summary>
        private static FamilySymbol GetIntersectFamilySymbol(Document doc)
        {
            FamilySymbol[] oldFamSymbOfGM = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .Where(el => el is FamilySymbol famSymb && famSymb.FamilyName == ClashPointFamilyName)
                .Cast<FamilySymbol>()
                .ToArray();

            // Если в проекте нет - то грузим
            if (!oldFamSymbOfGM.Any())
            {
                string path = $@"{Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)}\Source\RevitData\{_revitVersion}\{ClashPointFamilyName}.rfa";
                bool result = doc.LoadFamily(path);
                if (!result)
                    throw new Exception("Семейство для метки не найдено! Обратись к разработчику.");

                doc.Regenerate();

                // Повторяем после загрузки семейства
                oldFamSymbOfGM = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_GenericModel)
                    .Where(el => el is FamilySymbol famSymb && famSymb.FamilyName == ClashPointFamilyName)
                    .Cast<FamilySymbol>()
                    .ToArray();
            }

            FamilySymbol searchSymbol = oldFamSymbOfGM.FirstOrDefault();
            searchSymbol.Activate();

            return searchSymbol;
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
