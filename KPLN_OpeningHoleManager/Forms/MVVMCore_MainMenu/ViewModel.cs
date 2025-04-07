using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_OpeningHoleManager.Core;
using KPLN_OpeningHoleManager.ExecutableCommand;
using KPLN_OpeningHoleManager.Forms.MVVMCommand;
using KPLN_OpeningHoleManager.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;

namespace KPLN_OpeningHoleManager.Forms.MVVMCore_MainMenu
{
    internal class ViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Комманда: Выбрать задания и расстваить отверстия в модели
        /// </summary>
        public ICommand SetOpenHoleByTaskCommand { get; }

        public ViewModel()
        {
            SetOpenHoleByTaskCommand = new RelayCommand(CreateOpeningHole);
        }

        /// <summary>
        /// Реализация: Выбрать задания и расстваить отверстия в модели
        /// </summary>
        private void CreateOpeningHole()
        {
            try
            {
                if (Module.CurrentUIApplication == null) return;

                UIDocument uidoc = Module.CurrentUIApplication.ActiveUIDocument;
                Document doc = uidoc.Document;

                List<IOSOpeningHoleTaskEntity> iosTasks = GetIOSTasksFromLink(uidoc, doc);
                if (!iosTasks.Any())
                    return;

                List<AROpeningHoleEntity> arEntities = GetAROpeningsFromIOSTask(doc, iosTasks);

                if (arEntities.Any())
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new OpeningByIOSTaskMaker(arEntities));
            }
            catch (Exception ex)
            {
                new TaskDialog("Ошибка")
                {
                    MainIcon = TaskDialogIcon.TaskDialogIconError,
                    MainInstruction = $"Не удалось выполнить основную задачу. Отправь разработчику!\nОшибка: {ex.Message}",
                }.Show();
            }
        }

        /// <summary>
        /// Получить коллекцию заданий от ИОС
        /// </summary>
        private List<IOSOpeningHoleTaskEntity> GetIOSTasksFromLink(UIDocument uidoc, Document doc)
        {
            List<IOSOpeningHoleTaskEntity> iosTasks = new List<IOSOpeningHoleTaskEntity>();

            // Стварэнне фільтра для выбару толькі Mechanical Equipment у сувязях
            MechanicalEquipmentSelectionFilter selectionFilter = new MechanicalEquipmentSelectionFilter(doc);

            IList<Reference> pickedRefs;
            try
            {
                // Запуск выбару элементаў карыстальнікам
                pickedRefs = uidoc.Selection.PickObjects(
                    ObjectType.LinkedElement,
                    selectionFilter,
                    "Выделите задания от ИОС, которые нужно превартить в отверстия АР");
            }
            // Отмена пользователем
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return null; }

            foreach (Reference reference in pickedRefs)
            {
                RevitLinkInstance linkInstance = doc.GetElement(reference.ElementId) as RevitLinkInstance;
                if (linkInstance == null)
                    continue;

                Document linkedDoc = linkInstance.GetLinkDocument();
                if (linkedDoc == null)
                    continue;

                Element linkedElement = linkedDoc.GetElement(reference.LinkedElementId);
                if (linkedElement != null && linkedElement is FamilyInstance fi)
                {
                    LocationPoint locPnt = linkedElement.Location as LocationPoint;
                    IOSOpeningHoleTaskEntity iosTask = new IOSOpeningHoleTaskEntity(linkedDoc, linkedElement, locPnt.Point)
                        .SetFamilyPathAndName(linkedDoc)
                        .SetShapeByFamilyName(fi)
                        .SetTransform(linkInstance as Instance)
                        as IOSOpeningHoleTaskEntity;

                    iosTask.SetGeomParams();
                    iosTask.SetGeomParamsData();
                    iosTasks.Add(iosTask);
                }
            }

            return iosTasks;
        }

        /// <summary>
        /// Получить коллекцию заданий от ИОС
        /// </summary>
        private List<AROpeningHoleEntity> GetAROpeningsFromIOSTask(Document doc, List<IOSOpeningHoleTaskEntity> iosTasks)
        {
            List<AROpeningHoleEntity> arEntities = new List<AROpeningHoleEntity>();

            List<Element> arPotentialHosts = GetPotentialARHosts(doc, iosTasks);

            foreach (IOSOpeningHoleTaskEntity iosTask in iosTasks)
            {
                XYZ arDocCoord = iosTask.OHE_LinkTransform.OfPoint(iosTask.OHE_Point);

                Element hostElem = GetHostForOpening(arPotentialHosts, iosTask);
                if (hostElem == null)
                {
                    HtmlOutput.Print($"Для задания с id: {iosTask.OHE_Element.Id} из файла {iosTask.OHE_LinkDocument.Title} не удалось найти основу. Выполни расстановку вручную",
                        MessageType.Error);
                    continue;
                }

                AROpeningHoleEntity arEntity = new AROpeningHoleEntity(arDocCoord, hostElem, iosTask.OHE_Shape, MainDBService.Get_DBDocumentSubDepartment(iosTask.OHE_LinkDocument).Code)
                    .SetFamilyPathAndName(doc)
                    as AROpeningHoleEntity;

                arEntity.UpdatePointData(iosTask);
                arEntity.SetGeomParams();
                arEntity.SetGeomParamsData(iosTask.OHE_Height, iosTask.OHE_Width, iosTask.OHE_Radius);

                arEntities.Add(arEntity);
            }

            return arEntities;
        }

        /// <summary>
        /// Получить списко потенциальных основ для отверстия на основе выбранных заданий от ИОС
        /// </summary>
        private List<Element> GetPotentialARHosts(Document doc, List<IOSOpeningHoleTaskEntity> iosTasks)
        {
            List<Element> result = new List<Element>();

            BoundingBoxXYZ filterBBox = GeometryWorker.CreateOverallBBox(iosTasks);
            Outline filterOutline = GeometryWorker.CreateFilterOutline(filterBBox, 3);

            BoundingBoxIntersectsFilter bboxIntersectFilter = new BoundingBoxIntersectsFilter(filterOutline, 0.1);
            BoundingBoxIsInsideFilter bboxInsideFilter = new BoundingBoxIsInsideFilter(filterOutline, 0.1);

            // Коллекция ВСЕХ возможнных основ (лучше брать заново, т.к. кэш может протухнуть из-за правок модели)
            Element[] allHostFromDocumentColl = new FilteredElementCollector(doc)
                .OfClass(typeof(Wall))
                .ToArray();

            // Подготовка коллекции эл-в пересекаемых и внутри расширенного BoundingBoxXYZ
            result.AddRange(allHostFromDocumentColl
                .Where(e => bboxIntersectFilter.PassesFilter(doc, e.Id)));

            result.AddRange(allHostFromDocumentColl
                .Where(e => bboxInsideFilter.PassesFilter(doc, e.Id)));

            return result;
        }

        /// <summary>
        /// Получить стену, в которую будем ставить отверстие
        /// </summary>
        /// <returns></returns>
        private Element GetHostForOpening(List<Element> arPotentialHosts, IOSOpeningHoleTaskEntity iosTask)
        {
            Element host = null;
            double maxIntersection = 0;
            foreach (Element arPotHost in arPotentialHosts)
            {
                Solid arPotHostSolid = GeometryWorker.GetElemSolid(arPotHost);
                Solid intersectionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(arPotHostSolid, iosTask.OHE_Solid, BooleanOperationsType.Intersect);
                if (intersectionSolid != null && intersectionSolid.Volume > maxIntersection)
                {
                    maxIntersection = intersectionSolid.Volume;
                    host = arPotHost;
                }
            }

            return host;
        }

        private void OnPropertyChanged([CallerMemberName] string name = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    internal sealed class MechanicalEquipmentSelectionFilter : ISelectionFilter
    {
        private readonly Document _doc;
        private Document _linkDoc = null;

        public MechanicalEquipmentSelectionFilter(Document doc)
        {
            _doc = doc;
        }

        public bool AllowElement(Element elem) => true;

        public bool AllowReference(Reference reference, XYZ position)
        {
            if (_doc.GetElement(reference) is RevitLinkInstance rli)
                _linkDoc = rli.GetLinkDocument();
            else
            {
                _linkDoc = null;
                return false;
            }

            if (_linkDoc.GetElement(reference.LinkedElementId) is FamilyInstance fi)
                return fi.Category != null
                    && fi.Category.Id.IntegerValue == (int)BuiltInCategory.OST_MechanicalEquipment;

            return false;
        }
    }
}
