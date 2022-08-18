using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static KPLN_Loader.Output.Output;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    class CommandTagWiper : IExternalCommand
    {

        /// <summary>
        /// Список элементов, которые относятся к ошибкам
        /// </summary>
        private List<ElementId> _errorList = new List<ElementId>();

        /// <summary>
        /// Словарь элементов, где ключ - вид, значения - марки на виде
        /// </summary>
        private Dictionary<ElementId, List<ElementId>> _errorDict = new Dictionary<ElementId, List<ElementId>>();

        /// <summary>
        /// Список элементов, которые были исправлены
        /// </summary>
        private List<ElementId> _correctedList = new List<ElementId>();

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get application and documnet objects
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            View activeView = uidoc.ActiveView;

            // Обрабатываю пользовательскую выборку листов
            List<ViewSheet> sheetsList = new List<ViewSheet>();
            List<ElementId> selIds = uidoc.Selection.GetElementIds().ToList();
            if (selIds.Count > 0)
            {
                foreach (ElementId selId in selIds)
                {
                    Element elem = doc.GetElement(selId);
                    int catId = elem.Category.Id.IntegerValue;
                    if (catId.Equals((int)BuiltInCategory.OST_Sheets))
                    {
                        ViewSheet curViewSheet = elem as ViewSheet;
                        sheetsList.Add(curViewSheet);
                    }
                }
                if (sheetsList.Count == 0)
                {
                    TaskDialog.Show("Ошибка", "В выборке нет ни одного листа :(", TaskDialogCommonButtons.Ok);
                    return Result.Cancelled;
                }
            }

            // Поиск элементов на удаление
            try
            {
                // Анализирую выбранные листы
                if (sheetsList.Count > 0)
                {
                    foreach (ViewSheet viewSheet in sheetsList)
                    {
                        FindAllElementsOnList(doc, viewSheet);
                    }
                    ShowResult(doc);
                }

                // Анализирую все видовые экраны активного листа
                else if (activeView.Category.Id.IntegerValue.Equals((int)BuiltInCategory.OST_Sheets))
                {
                    ViewSheet viewSheet = activeView as ViewSheet;
                    FindAllElementsOnList(doc, viewSheet);
                    ShowResult(doc);
                }

                // Анализирую вид
                else
                {
                    FindAllElements(doc, activeView.Id);
                    ShowResult(doc);
                }
            }
            catch (Exception ex)
            {
                PrintError(ex);
            }
            return Result.Succeeded;
        }

        /// <summary>
        /// Метод для поиска в модели элементов аннотаций на единице выбранного элемента и записи в коллекцию
        /// </summary>
        private void FindAllElements(Document doc, ElementId viewId)
        {
            List<ElementId> errorTags = new List<ElementId>();

            ICollection<ElementId> collection = new FilteredElementCollector(doc, viewId).OfCategory(BuiltInCategory.OST_RoomTags).WhereElementIsNotElementType().ToElementIds();
                
            
            foreach (ElementId elementId in collection)
            {
                RoomTag roomTag = doc.GetElement(elementId) as RoomTag;
                if (roomTag.TaggedRoomId.LinkedElementId.IntegerValue == -1)
                {
                    errorTags.Add(elementId);
                }
                
            }

            if (errorTags.Count > 0)
            {
                if (_errorDict.ContainsKey(viewId))
                {
                    _errorDict[viewId].AddRange(errorTags);
                }
                else
                {
                    _errorDict.Add(viewId, errorTags);
                }
            }
        }

        /// <summary>
        /// Метод для поиска в модели элементов аннотаций на листах и записи в коллекцию или словарь (в зависимости от количества выбранных листов)
        /// </summary>
        private void FindAllElementsOnList(Document doc, ViewSheet viewSheet)
        {
            // Анализирую размещенные виды
            ICollection<ElementId> allViewPorts = viewSheet.GetAllViewports();
            foreach (ElementId vpId in allViewPorts)
            {
                Viewport vp = (Viewport)doc.GetElement(vpId);
                ElementId viewId = vp.ViewId;
                Element currentElement = doc.GetElement(viewId);
                
                // Анализирую только виды
                if (currentElement.GetType().Equals(typeof(ViewPlan)) || currentElement.GetType().Equals(typeof(ViewSection)))
                {
                    FindAllElements(doc, viewId);
                }
            }
        }

        /// <summary>
        /// Метод для нахождения марке новой основы и создания новой марки вместо старой
        /// </summary>
        private void CorrectedElements(Document doc)
        {
            ICollection<ElementId> collection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType().ToElementIds();
            foreach (ElementId linktId in collection)
            {
                RevitLinkInstance linkInst = doc.GetElement(linktId) as RevitLinkInstance;
                
                if (linkInst.Name.ToLower().Contains("ar") || linkInst.Name.ToLower().Contains("ар"))
                {
                    Document linkDoc = linkInst.GetLinkDocument();
                    
                    if (linkDoc != null) 
                    { 
                        FilteredElementCollector roomsColl = new FilteredElementCollector(linkDoc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType();
                        if (roomsColl.Count() > 0)
                        {
                            //var transformCoord = linkInst.GetTransform();

                            foreach (KeyValuePair<ElementId, List<ElementId>> kvp in _errorDict)
                            {
                                foreach (ElementId errId in kvp.Value)
                                {
                                    RoomTag roomTag = doc.GetElement(errId) as RoomTag;
                                    LocationPoint tagLocationPoint = roomTag.Location as LocationPoint;
                                    XYZ tagPoint = tagLocationPoint.Point;

                                    foreach (Element elem in roomsColl)
                                    {
                                        //transformCoord.OfPoint(tagPoint);
                                        Room room = elem as Room;
                                        if (room.IsPointInRoom(tagPoint))
                                        {
                                            //Print($"{linkInst.Name}/{room.Name}", KPLN_Loader.Preferences.MessageType.Regular);
                                            LinkElementId roomId = new LinkElementId(linkInst.Id, room.Id);
                                            var uvPoint = new UV(tagPoint.X, tagPoint.Y);
                                            
                                            var newRoomTag = doc.Create.NewRoomTag(roomId, uvPoint, kvp.Key);
                                            newRoomTag.RoomTagType = roomTag.RoomTagType;


                                            _correctedList.Add(errId);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Метод для вывода результатов пользователю, а также для удаления элементов в модели
        /// </summary>
        private void ShowResult(Document doc)
        {
            // Обрабатываю элементы в модели
            using (Transaction t = new Transaction(doc))
            {
                t.Start("KPLN_Почистить/починить марки помещений");
                
                CorrectedElements(doc);
                foreach (KeyValuePair<ElementId, List<ElementId>> kvp in _errorDict)
                {
                    doc.Delete(kvp.Value);
                }
                
                t.Commit();
            }

            // Вывожу результат пользователю
            if (_errorDict.Count == 0)
            {
                TaskDialog.Show("Результат", "Элементы не обнаружены :)", TaskDialogCommonButtons.Ok);
            }
            else if (_correctedList.Count != 0)
            {
                TaskDialog.Show("Результат исправления", $"Элементы исправлены. Количество - {_correctedList.Count}", TaskDialogCommonButtons.Ok);
            }
            else if (_errorDict.Count != 0)
            {
                TaskDialog.Show("Результат удаления", $"Элементы удалены. Количество - {_errorDict.Values.Count}", TaskDialogCommonButtons.Ok);
            }

        }
    }
}
