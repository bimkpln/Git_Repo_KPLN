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
    /// Класс для ОЧИСТКИ элементов пересечений при удалении эл-в
    /// </summary>
    internal sealed class IntersectPointDelElemCleaner : IntersectPointMaker, IExecutableCommand
    {
        /// <summary>
        /// Имя транзакции для анализа на наличие
        /// </summary>
        public static readonly string TransName = "KPLN: Очистка меток коллизий";
        
        private readonly ElementId[] _deletedElemIds;

        public IntersectPointDelElemCleaner(IEnumerable<ElementId> deletedElemIDs)
        {
            _deletedElemIds = deletedElemIDs.ToArray();
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

            if (_deletedElemIds.Count() == 0)
                return Result.Cancelled;

            Element[] oldPointElems = GetOldIntersectionPointsEntities(doc);
            if(oldPointElems.Length == 0)
                return Result.Cancelled;

            using (Transaction trans = new Transaction(doc, TransName))
            {
                trans.Start();

                // Удаляю существующте по удаленным элементам
                List<Element> deleteElemToDel = GetIntersectionPointsEntitiesToDelete(doc, oldPointElems, _deletedElemIds.Select(id => id.IntegerValue));
                if (deleteElemToDel != null)
                    doc.Delete(deleteElemToDel.Select(el => el.Id).ToArray());

                trans.Commit();
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Получить коллекцию сущностей точек пересечения в проекте для удаленных элементов ДЛЯ УДАЛЕНИЯ
        /// </summary>
        private static List<Element> GetIntersectionPointsEntitiesToDelete(Document doc, IEnumerable<Element> pointElemsInModel, IEnumerable<int> deletedElemIds)
        {
            List<Element> resultToDel = new List<Element>();

            List<IntersectPointEntity> oldEntitiesInModel = new List<IntersectPointEntity>();
            foreach (Element pointElem in pointElemsInModel)
            {
                // Получаю данные об элементах
                int addedElemIdInModel = -1;
                int oldElemIdInModel = -1;
                try
                {
                    addedElemIdInModel = int.Parse(pointElem.get_Parameter(AddedElementId_Param).AsString());
                    oldElemIdInModel = int.Parse(pointElem.get_Parameter(OldElementId_Param).AsString());
                }
                catch
                {
                    throw new FormatException($"Отправь разработчику: Не удалось конвертировать данные из id-элементов для эл-та с id:{pointElem.Id}");
                }

                // Анализ наличия коллизии (если эл-т в списке на удаление)
                if (deletedElemIds.Any(cId => cId == addedElemIdInModel || cId == oldElemIdInModel))
                    resultToDel.Add(pointElem);
            }

            return resultToDel;
        }
    }
}
