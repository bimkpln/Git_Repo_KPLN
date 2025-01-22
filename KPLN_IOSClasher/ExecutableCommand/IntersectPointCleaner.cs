using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_IOSClasher.Core;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_IOSClasher.ExecutableCommand
{
    /// <summary>
    /// Класс для ОЧИСТКИ элементов пересечений
    /// </summary>
    internal class IntersectPointCleaner : IntersectPointFamInst, IExecutableCommand
    {
        /// <summary>
        /// Имя транзакции для анализа на наличие
        /// </summary>
        public static readonly string TransName = "KPLN: Очистка меток коллизий";
        
        private readonly ElementId[] _deletedElems;

        public IntersectPointCleaner(IEnumerable<ElementId> deletedElemIDs)
        {
            _deletedElems = deletedElemIDs.ToArray();
        }

        public Result Execute(UIApplication app)
        {
            // Такое возможно при работе плагинов с открываением/сохранением моделей (модель не открылась)
            if (app.ActiveUIDocument == null)
                return Result.Cancelled;

            Document doc = app.ActiveUIDocument.Document;
            UIDocument uidoc = app.ActiveUIDocument;

            if (uidoc == null)
                return Result.Cancelled;

            using (Transaction trans = new Transaction(doc, TransName))
            {
                trans.Start();

                // Удаляю существующте по удаленным элементам
                List<Element> deleteElemToDel = GetIntersectionPointsEntitiesToDelete(doc);
                if (deleteElemToDel != null)
                    doc.Delete(deleteElemToDel.Select(el => el.Id).ToArray());

                trans.Commit();
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Получить коллекцию сущностей точек пересечения в проекте для удаленных элементов ДЛЯ УДАЛЕНИЯ
        /// </summary>
        private List<Element> GetIntersectionPointsEntitiesToDelete(Document doc)
        {
            List<Element> resultToDel = new List<Element>();

            Element[] oldPointElemsInModel = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .Where(el => el is FamilyInstance famInst && famInst.Symbol.FamilyName == ClashPointFamilyName)
                .ToArray();

            if (oldPointElemsInModel.Length == 0)
                return null;

            List<IntersectPointEntity> oldEntitiesInModel = new List<IntersectPointEntity>(oldPointElemsInModel.Length);
            foreach (Element oldElem in oldPointElemsInModel)
            {
                if (int.TryParse(oldElem.get_Parameter(AddedElementId_Param).AsString(), out int oldAddedElemId))
                {
                    // Если старый элемент, НИКАК не упоминается в коллекции добавленных элементы, то игнорируем
                    if (_deletedElems.All(id => id.IntegerValue != oldAddedElemId))
                        continue;

                    resultToDel.Add(oldElem);
                }
                else
                    throw new FormatException($"Отправь разработчику: Не удалось конвертировать данные из id-элемента для эл-та с id:{oldElem.Id}");

            }

            return resultToDel;
        }
    }
}
