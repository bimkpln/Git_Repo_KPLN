using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_IOSClasher.Core;
using KPLN_IOSClasher.Services;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_SQLiteWorker;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace KPLN_IOSClasher.ExecutableCommand
{
    /// <summary>
    /// Класс для АНАЛИЗА пересечений с участием СВЯЗЕЙ
    /// </summary>
    internal sealed class IntersectPointLinkWorker : IntersectPointMaker, IExecutableCommand
    {
        /// <summary>
        /// Имя транзакции для анализа на наличие
        /// </summary>
        public static readonly string TransName = "KPLN: Обновление меток коллизий (связи)";

        public IntersectPointLinkWorker()
        {
            CultureInfo.CurrentCulture = new CultureInfo("ru-RU");
        }

        public Result Execute(UIApplication app)
        {
            // Такое возможно при работе плагинов с открываением/сохранением моделей (модель не открылась)
            if (app.ActiveUIDocument == null)
                return Result.Cancelled;

            Document doc = app.ActiveUIDocument.Document;
            UIDocument uidoc = app.ActiveUIDocument;

            if (uidoc == null) return Result.Cancelled;

            using (Transaction trans = new Transaction(doc, TransName))
            {
                trans.Start();

                Element[] oldPointElems = GetOldIntersectionPointsEntities(doc);
                if (oldPointElems.Length == 0) return Result.Cancelled;

                // Получаю данные по точкам
                Dictionary<ElementId, IntersectPointEntity> oldLinkPointEntities = CreateIntPntEntities_ByOldPoints(doc, oldPointElems);

                // Удаляю не актуальные (если можно их удалить). Проблема с занятами клэшпоинтами приводит к ложным клэшам. В пределах погрешности ок, ведь когда 
                // юзер зайдет в модель - он свои клэши почистит (для него эти эл-ты уже не заняты)
                ICollection<ElementId> oldElemsToDel = GetOldToDelete_ByOldPoints(doc, oldLinkPointEntities);
                ICollection<ElementId> availableWSElemsId = WorksharingUtils.CheckoutElements(doc, oldElemsToDel);
                doc.Delete(availableWSElemsId);

                foreach (ElementId delElId in oldElemsToDel)
                {
                    oldLinkPointEntities.Remove(delElId);
                }

                // Создаю новые
                IntersectPointEntity[] clearedPointEntities = ClearedNeighbourEntities(oldPointElems, oldLinkPointEntities.Values);
                CreateIntersectFamilyInstance(doc, clearedPointEntities);

                trans.Commit();
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Получить коллекцию IntersectElemDocEntity по СУЩЕСТВУЩИМ точкам пересечений для ЛИНКОВ
        /// </summary>
        private static Dictionary<ElementId, IntersectPointEntity> CreateIntPntEntities_ByOldPoints(Document doc, Element[] pointElemsInModel)
        {
            // ElementId - это клэшпоинт ИЗ модели, а IntersectPointEntity - это экземпляр класса сущности, по данному клэшпоинту из модели
            Dictionary<ElementId, IntersectPointEntity> result = new Dictionary<ElementId, IntersectPointEntity>();
            foreach (Element pointElem in pointElemsInModel)
            {
                // Проверка на наличие элементов коллизий с участием линков (внутри дока анализируются не тут)
                string linkParamData = pointElem.get_Parameter(LinkInstanceId_Param).AsString();
                if (string.IsNullOrEmpty(linkParamData) || linkParamData.Equals("-1")) continue;

                // Получаю данные об элементах
                int firstElemId = -1;
                int secondElemId = -1;
                try
                {
                    firstElemId = int.Parse(pointElem.get_Parameter(AddedElementId_Param).AsString());
                    secondElemId = int.Parse(pointElem.get_Parameter(OldElementId_Param).AsString());
                }
                catch
                {
                    throw new FormatException($"Отправь разработчику: Не удалось конвертировать данные из id-элементов для эл-та с id:{pointElem.Id}");
                }

                // Получаю данные по элементу №2 (из линка)
                if (int.TryParse(linkParamData, out int linkId))
                {
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                    int linkIdValue = linkId;
#else
                    long linkIdValue = (long)linkId;
#endif
                    Element linkElemInst = doc.GetElement(new ElementId(linkIdValue));
                    if (linkElemInst is RevitLinkInstance linkInst)
                    {
                        XYZ oldPoint = ParseStringToXYZ(pointElem.get_Parameter(PointCoord_Param).AsString());
                        result.Add(pointElem.Id, new IntersectPointEntity(oldPoint, firstElemId, secondElemId, linkId, DBMainService.CurrentDBUser));
                    }
                    else
                        throw new FormatException($"Отправь разработчику: Не удалось конвертировать id-связь в RevitLinkInstance для эл-та с id:{pointElem.Id}");
                }
                else
                    throw new FormatException($"Отправь разработчику: Не удалось конвертировать данные из id-связи для эл-та с id:{pointElem.Id}");
            }

            return result;
        }

        /// <summary>
        /// Подготовка коллекции на удаление из размещенных ранее клэшпоинтов
        /// </summary>
        private static List<ElementId> GetOldToDelete_ByOldPoints(Document doc, Dictionary<ElementId, IntersectPointEntity> oldPointEntitiesDict)
        {
            List<ElementId> resultToDel = new List<ElementId>();

            foreach (KeyValuePair<ElementId, IntersectPointEntity> kvp in oldPointEntitiesDict)
            {
                Element oldPntElem = doc.GetElement(kvp.Key);
                IntersectPointEntity oldPntEntity = kvp.Value;


                // Проверка файла на наличие элемента (должны чиститься, но вполне могут Redo/Undo не до конца прокликать).
                // Если эл-та нет - удаляем
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                Element addedElemInModel = doc.GetElement(new ElementId((int)oldPntEntity.AddedElement_Id));
#else
                Element addedElemInModel = doc.GetElement(new ElementId(oldPntEntity.AddedElement_Id));
#endif
                if (addedElemInModel == null)
                {
                    resultToDel.Add(kvp.Key);
                    continue;
                }


                // Проверка линка на наличие элемента (должны чиститься, но вполне могут Redo/Undo не до конца прокликать).
                // Если эл-та нет - удаляем
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                RevitLinkInstance linkInst = doc.GetElement(new ElementId((int)oldPntEntity.LinkInstance_Id)) as RevitLinkInstance;
#else
                RevitLinkInstance linkInst = doc.GetElement(new ElementId(oldPntEntity.LinkInstance_Id)) as RevitLinkInstance;
#endif
                Document linkDoc = linkInst.GetLinkDocument();

                // Если док не подгружен - linkDoc не взять. Просто игнор, до момента подгрузки
                if (linkDoc == null) continue;

#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                Element oldElemInMode = linkDoc.GetElement(new ElementId((int)oldPntEntity.OldElement_Id));
#else
                Element oldElemInMode = linkDoc.GetElement(new ElementId(oldPntEntity.OldElement_Id));
#endif
                if (oldElemInMode == null)
                {
                    resultToDel.Add(kvp.Key);
                    continue;
                }


                // Анализ наличия коллизии
                Solid addedElemSolid = DocController.GetElemSolid(addedElemInModel);
                if (addedElemSolid == null)
                {
                    HtmlOutput.Print($"У элемента с id: {addedElemInModel.Id} из ТВОЕЙ модели проблемы с получением Solid. " +
                        $"Элемент проигнорирован, ошибку стоит отправить разработчику);", MessageType.Error);
                    continue;
                }

                XYZ oldPoint = ParseStringToXYZ(oldPntElem.get_Parameter(PointCoord_Param).AsString());

                BoundingBoxXYZ pntEntBBox = DocController.CreateElemsBBox(new List<Element>(1) { oldPntElem });
                Outline filterOutline = DocController.CreateFilterOutline(pntEntBBox, 1);

                IntersectCheckEntity oldDocCheckEnt = new IntersectCheckEntity(doc, filterOutline, linkInst);

                IEnumerable<IntersectPointEntity> docIntPntEntities = oldDocCheckEnt
                    .CurrentDocElemsToCheck
                    .Select(ent => DocController.GetPntEntityFromElems(addedElemInModel, addedElemSolid, ent, oldDocCheckEnt));

                // Если новые IntersectPointEntity не сгенрировались - коллизия для дока ушла
                if (docIntPntEntities.All(elemToCheck => elemToCheck == null))
                {
                    resultToDel.Add(kvp.Key);
                    continue;
                }

                foreach (IntersectPointEntity docEntToCheck in docIntPntEntities)
                {
                    // Анализ на совпадение по id и по координатам
                    if (docEntToCheck != null && oldPntEntity.AddedElement_Id == docEntToCheck.AddedElement_Id && oldPntEntity.OldElement_Id == docEntToCheck.OldElement_Id)
                    {
                        // Анализ на совпадение по координатам, если его нет ДЛЯ ВСЕХ из коллекции - удаляем
                        if (!docIntPntEntities.Any(ent => Math.Abs(docEntToCheck.IntersectPoint.DistanceTo(oldPoint)) < 0.05))
                            resultToDel.Add(kvp.Key);
                    }
                }
            }

            return resultToDel;
        }
    }
}
