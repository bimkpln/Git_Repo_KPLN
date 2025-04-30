using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_IOSClasher.Core;
using KPLN_IOSClasher.Services;
using KPLN_Library_Forms.UI.HtmlWindow;
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
    internal sealed class IntersectPointDocWorker : IntersectPointMaker, IExecutableCommand
    {
        /// <summary>
        /// Имя транзакции для анализа на наличие
        /// </summary>
        public static readonly string TransName = "KPLN: Создание меток коллизий";

        private readonly IntersectPointEntity[] _intersectPointEntities;
        private readonly Element[] _checkedElems;

        public IntersectPointDocWorker(IEnumerable<IntersectPointEntity> intersectPoints, IEnumerable<Element> checkedElems)
        {
            CultureInfo.CurrentCulture = new CultureInfo("ru-RU");

            _intersectPointEntities = intersectPoints.ToArray();
            _checkedElems = checkedElems.ToArray();
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

                Element[] oldPointElems = GetOldIntersectionPointsEntities(doc);
                IntersectPointEntity[] clearedPointEntities = null;
                
                // Удаляю не актуальные
                if (oldPointElems.Length != 0)
                {
                    // Удаляю не актуальные (если можно их удалить). Проблема с занятами клэшпоинтами приводит к ложным клэшам. В пределах погрешности ок, ведь когда 
                    // юзер зайдет в модель - он свои клэши почистит (для него эти эл-ты уже не заняты)
                    List<Element> oldElemsToDel = GetOldToDelete_ByOldPointsAndChackedElems(doc, oldPointElems, _checkedElems.Select(el => el.Id.IntegerValue));
                    ICollection<ElementId> availableWSElemsId = WorksharingUtils.CheckoutElements(doc, oldElemsToDel.Select(el => el.Id).ToArray());
                    doc.Delete(availableWSElemsId);

                    clearedPointEntities = ClearedNeighbourEntities(oldPointElems, _intersectPointEntities);
                }
                else
                    clearedPointEntities = _intersectPointEntities;

                // Создаю новые
                CreateIntersectFamilyInstance(doc, clearedPointEntities);

                trans.Commit();

                if (clearedPointEntities.Any())
                {
                    if (!KPLN_Loader.Application.OnIdling_CommandQueue.Any(i => i is IntersectPointDocWorker))
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
        /// Подготовка коллекции на удаление из размещенных ранее клэшпоинтов
        /// </summary>
        private static List<Element> GetOldToDelete_ByOldPointsAndChackedElems(Document doc, IEnumerable<Element> pointElemsInModel, IEnumerable<int> checkedElemIds)
        {
            List<Element> resultToDel = new List<Element>();

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

                // Анализ наличия коллизии (если эл-т в списке на редактирование)
                if (checkedElemIds.Any(cId => cId == addedElemIdInModel || cId == oldElemIdInModel))
                {
                    // Проверка файла на наличие элемента (должны чиститься, но вполне могут Redo/Undo не до конца прокликать).
                    // Если эл-та нет - удаляем
                    Element addedElemInModel = doc.GetElement(new ElementId(addedElemIdInModel));
                    if (addedElemInModel == null)
                    {
                        resultToDel.Add(addedElemInModel);
                        continue;
                    }


                    // Проверка файла или линка (в зависимости от эл-та) на наличие элемента (должны чиститься, но вполне могут Redo/Undo не до конца прокликать).
                    // Если эл-та нет - удаляем
                    Element oldElemInModel = null;
                    string linkInstInModel = pointElem.get_Parameter(LinkInstanceId_Param).AsString();
                    RevitLinkInstance linkInst = null;
                    Transform linkTrans = null;
                    if (!string.IsNullOrEmpty(linkInstInModel) && linkInstInModel != "-1")
                    {
                        if (int.TryParse(linkInstInModel, out int linkId))
                        {
                            linkInst = doc.GetElement(new ElementId(linkId)) as RevitLinkInstance;
                            if (linkInst == null)
                                throw new FormatException($"Отправь разработчику: Не удалось конвертировать id-связь в RevitLinkInstance для эл-та с id:{pointElem.Id}");
                            else
                            {
                                Document linkDoc = linkInst.GetLinkDocument();
                                oldElemInModel = linkDoc.GetElement(new ElementId(oldElemIdInModel));
                                linkTrans = IntersectCheckEntity.GetLinkTransform(linkInst);
                            }
                            
                        }
                        else
                            throw new FormatException($"Отправь разработчику: Не удалось конвертировать данные из id-связи для эл-та с id:{pointElem.Id}");
                    }
                    else
                        oldElemInModel = doc.GetElement(new ElementId(oldElemIdInModel));

                    if (oldElemInModel == null)
                    {
                        resultToDel.Add(pointElem);
                        continue;
                    }


                    // Анализ наличия коллизии между эл-тами
                    Solid addedElemSolidInModel = DocController.GetElemSolid(addedElemInModel);
                    if (addedElemSolidInModel == null)
                    {
                        HtmlOutput.Print($"У элемента с id: {addedElemInModel.Id} из ТВОЕЙ модели проблемы с получением Solid. " +
                            $"Элемент проигнорирован, ошибку стоит отправить разработчику);", MessageType.Error);
                        continue;
                    }

                    Solid oldElemSolidInModel = DocController.GetElemSolid(oldElemInModel, linkTrans);
                    if (oldElemSolidInModel == null)
                    {
                        string varStr = linkInst != null ? $"СВЯЗИ с id: {linkInst.Id}" : "ТВОЕЙ модели";
                        HtmlOutput.Print($"У элемента с id: {oldElemInModel.Id} из {varStr} проблемы с получением Solid. " +
                            $"Элемент проигнорирован, ошибку стоит отправить разработчику);", MessageType.Error);
                        continue;
                    }

                    Solid intersectionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(addedElemSolidInModel, oldElemSolidInModel, BooleanOperationsType.Intersect);
                    if (intersectionSolid != null && intersectionSolid.Volume > 0)
                    {
                        XYZ pointElemXYZ = ParseStringToXYZ(pointElem.get_Parameter(PointCoord_Param).AsString());
                        XYZ intersectionXYZ = intersectionSolid.ComputeCentroid();
                        if (pointElemXYZ.DistanceTo(intersectionXYZ) < 0.05)
                            continue;
                    }

                    resultToDel.Add(pointElem);
                }
            }

            return resultToDel;
        }
    }
}
