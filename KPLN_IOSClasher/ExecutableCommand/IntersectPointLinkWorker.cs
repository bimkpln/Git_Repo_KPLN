using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_IOSClasher.Core;
using KPLN_IOSClasher.Services;
using KPLN_Library_Forms.UI.HtmlWindow;
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
    internal class IntersectPointLinkWorker : IntersectPointFamInst, IExecutableCommand
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

                Element[] oldPointElems = IntersectPointMaker.GetOldIntersectionPointsEntities(doc);
                if (oldPointElems.Length == 0) return Result.Cancelled;

                // Обновляю данные по точкам
                Dictionary<ElementId, IntersectPointEntity> oldLinkPointEntities = CreateIntPntEntities_ByOldPoints(doc, oldPointElems);

                // Удаляю не актуальные (если можно их удалить). Проблема с занятами клэшпоинтами приводит к ложным клэшам. В пределах погрешности ок, ведь когда 
                // юзер зайдет в модель - он свои клэши почистит (для него эти эл-ты уже не заняты)
                ICollection<ElementId> oldElemsToDel = GetOldToDelete_ByOldPoints(doc, oldLinkPointEntities, oldPointElems);
                ICollection<ElementId> availableWSElemsId = WorksharingUtils.CheckoutElements(doc, oldElemsToDel);
                doc.Delete(availableWSElemsId);

                foreach (ElementId delElId in oldElemsToDel)
                {
                    oldLinkPointEntities.Remove(delElId);
                }

                // Создаю новые
                IntersectPointEntity[] clearedPointEntities = IntersectPointMaker.ClearedNewEntities(oldPointElems, oldLinkPointEntities.Values);
                IntersectPointMaker.CreateIntersectFamilyInstance(doc, clearedPointEntities);

                trans.Commit();
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Получить коллекцию IntersectElemDocEntity по СУЩЕСТВУЩИМ точкам пересечений для ЛИНКОВ
        /// </summary>
        private static Dictionary<ElementId, IntersectPointEntity> CreateIntPntEntities_ByOldPoints(Document doc, Element[] pointElemsInModel)
        {
            Dictionary<ElementId, IntersectPointEntity> result = new Dictionary<ElementId, IntersectPointEntity>();
            foreach (Element pointElem in pointElemsInModel)
            {
                // Проверка на наличие элементов коллизий с участием линков (внутри дока анализируются не тут)
                string linkParamData = pointElem.get_Parameter(LinkInstanceId_Param).AsString();
                if (string.IsNullOrEmpty(linkParamData) || linkParamData.Equals("-1")) continue;

                // Получаю данные об элементах
                string firstElemParamData = pointElem.get_Parameter(AddedElementId_Param).AsString();
                string secondElemParamData = pointElem.get_Parameter(OldElementId_Param).AsString();
                int firstElemId = -1;
                int secondElemId = -1;
                try
                {
                    firstElemId = int.Parse(firstElemParamData);
                    secondElemId = int.Parse(secondElemParamData);
                }
                catch
                {
                    throw new FormatException($"Отправь разработчику: Не удалось конвертировать данные из id-элементов для эл-та с id:{pointElem.Id}");
                }

                // Получаю данные по элементу №2 (из линка)
                if (int.TryParse(linkParamData, out int linkId))
                {
                    Element linkElemInst = doc.GetElement(new ElementId(linkId));
                    if (linkElemInst is RevitLinkInstance linkInst)
                    {
                        XYZ oldPoint = ParseStringToXYZ(pointElem.get_Parameter(PointCoord_Param).AsString());
                        result.Add(pointElem.Id, new IntersectPointEntity(oldPoint, firstElemId, secondElemId, linkId, Module.ModuleDBWorkerService.CurrentDBUser));
                    }
                    else
                        throw new FormatException($"Отправь разработчику: Не удалось конвертировать id-связь в RevitLinkInstance для эл-та с id:{pointElem.Id}");
                }
                else
                    throw new FormatException($"Отправь разработчику: Не удалось конвертировать данные из id-связи для эл-та с id:{pointElem.Id}");
            }

            return result;
        }

        private static List<ElementId> GetOldToDelete_ByOldPoints(Document doc, Dictionary<ElementId, IntersectPointEntity> oldPointEntitiesDict, Element[] oldPointElems)
        {
            List<ElementId> resultToDel = new List<ElementId>();

            foreach (KeyValuePair<ElementId, IntersectPointEntity> kvp in oldPointEntitiesDict)
            {
                Element oldPntElem = doc.GetElement(kvp.Key);
                IntersectPointEntity oldPntEntity = kvp.Value;

                RevitLinkInstance linkInst = doc.GetElement(new ElementId(oldPntEntity.LinkInstance_Id)) as RevitLinkInstance;
                Document linkDoc = linkInst.GetLinkDocument();

                // Если док не подгружен - linkDoc не взять. Просто игнор, до момента подгрузки
                if (linkDoc == null) continue;

                // Проверка линка на наличие элемента в модели (если нет - удаляем)
                if (linkDoc.GetElement(new ElementId(oldPntEntity.OldElement_Id)) == null)
                {
                    resultToDel.Add(kvp.Key);
                    continue;
                }

                Element oldAddedElem = doc.GetElement(new ElementId(oldPntEntity.AddedElement_Id));
                // Проверка файла на наличие элемента (должны чиститься, но вполне могут Redo/Undo не до конца прокликать).Если эл-та нет - удаляем
                if (oldAddedElem == null)
                {
                    resultToDel.Add(kvp.Key);
                    continue;
                }

                // Проверка на соответствие полученного по id элемента - линейному эл-ту
                // (лёгкая компенсация промаха по id, если эл-т уалили, а потом новый создали с тем же id)
                if (IntersectCheckEntity.BuiltInCategories.Count(bic => (int)bic == oldAddedElem.Category.Id.IntegerValue) == 0)
                {
                    resultToDel.Add(kvp.Key);
                    continue;
                }


                Solid addedElemSolid = DocController.GetElemSolid(oldAddedElem);
                if (addedElemSolid == null)
                {
                    HtmlOutput.Print($"У элемента с id: {oldAddedElem.Id} из ТВОЕЙ модели проблемы с получением Solid. " +
                        $"Элемент проигнорирован, ошибку стоит отправить разработчику);", MessageType.Error);
                    continue;
                }

                XYZ oldPoint = ParseStringToXYZ(oldPntElem.get_Parameter(PointCoord_Param).AsString());

                BoundingBoxXYZ pntEntBBox = DocController.CreateElemsBBox(new List<Element>(1) { oldPntElem });
                Outline filterOutline = DocController.CreateFilterOutline(pntEntBBox, 1);

                IntersectCheckEntity oldDockEnt = new IntersectCheckEntity(doc, pntEntBBox, filterOutline, linkInst);

                IEnumerable<IntersectPointEntity> docIntPntEntities = oldDockEnt
                    .CurrentDocElemsToCheck
                    .Select(ent => DocController.GetPntEntityFromElems(oldAddedElem, addedElemSolid, ent, oldDockEnt));

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
