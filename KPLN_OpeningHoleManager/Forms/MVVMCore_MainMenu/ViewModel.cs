using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_OpeningHoleManager.Core;
using KPLN_OpeningHoleManager.ExecutableCommand;
using KPLN_OpeningHoleManager.Forms.MVVMCommand;
using KPLN_OpeningHoleManager.Services;
using Microsoft.SqlServer.Server;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_OpeningHoleManager.Forms.MVVMCore_MainMenu
{
    internal class ViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private double _openHoleExpandedValue = 50;

        /// <summary>
        /// Комманда: Расставить отверстия по элементам ИОС
        /// </summary>
        public ICommand CreateOpenHoleByIOSElemsCommand { get; }

        /// <summary>
        /// Комманда: Выбрать задания и расстваить отверстия в модели
        /// </summary>
        public ICommand SetOpenHoleByTaskCommand { get; }

        /// <summary>
        /// Комманда: Объеденить отверстия
        /// </summary>
        public ICommand UnitOpenHolesCommand { get; }

        /// <summary>
        /// Значение расширение отверстия при расстановке
        /// </summary>
        public double OpenHoleExpandedValue 
        { 
            get => _openHoleExpandedValue;
            set
            {
                _openHoleExpandedValue = value;
                OnPropertyChanged();
            }
        }

        public ViewModel()
        {
            CreateOpenHoleByIOSElemsCommand = new RelayCommand(CreateOpenHoleByIOSElems);
            SetOpenHoleByTaskCommand = new RelayCommand(SetOpenHoleByTask);
            UnitOpenHolesCommand = new RelayCommand(UnitOpenHoles);
        }

        #region Расставить отверстия по элементам ИОС
        /// <summary>
        /// Реализация: Расставить отверстия по элементам ИОС
        /// </summary>
        private void CreateOpenHoleByIOSElems()
        {

        }
        #endregion

        #region Выбрать задания и расстваить отверстия в модели
        /// <summary>
        /// Реализация: Выбрать задания и расстваить отверстия в модели
        /// </summary>
        private void SetOpenHoleByTask()
        {
            try
            {
                if (Module.CurrentUIApplication == null) return;

                UIDocument uidoc = Module.CurrentUIApplication.ActiveUIDocument;
                Document doc = uidoc.Document;

                List<IOSOpeningHoleTaskEntity> iosTasks = SelectAndGetIOSTasksFromLink(uidoc, doc);
                if (!iosTasks.Any())
                    return;

                List<AROpeningHoleEntity> arEntities = GetAROpeningsFromIOSTask(doc, iosTasks);

                if (arEntities.Any())
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new AR_OHE_Maker(arEntities));
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
        /// Выбрать коллекцию заданий от ИОС
        /// </summary>
        private List<IOSOpeningHoleTaskEntity> SelectAndGetIOSTasksFromLink(UIDocument uidoc, Document doc)
        {
            List<IOSOpeningHoleTaskEntity> iosTasks = new List<IOSOpeningHoleTaskEntity>();

            MechEquipLinkSelectionFilter selectionFilter = new MechEquipLinkSelectionFilter(doc);

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
        /// Получить коллекцию отверстий АР из заданий от ИОС
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

                AROpeningHoleEntity arEntity = new AROpeningHoleEntity(iosTask.OHE_Shape, MainDBService.Get_DBDocumentSubDepartment(iosTask.OHE_LinkDocument).Code, hostElem, arDocCoord)
                    .SetFamilyPathAndName(doc)
                    as AROpeningHoleEntity;

                arEntity.UpdatePointData(iosTask);
                arEntity.SetGeomParams();
                arEntity.SetGeomParamsRoundData(iosTask.OHE_Height, iosTask.OHE_Width, iosTask.OHE_Radius);

                arEntities.Add(arEntity);
            }

            return arEntities;
        }

        /// <summary>
        /// Получить список потенциальных основ для отверстия на основе выбранных заданий от ИОС
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
        #endregion

        #region Объеденить отверстия
        /// <summary>
        /// Реализация: Объеденить отверстия
        /// </summary>
        private void UnitOpenHoles()
        {
            try
            {
                if (Module.CurrentUIApplication == null) return;

                UIDocument uidoc = Module.CurrentUIApplication.ActiveUIDocument;
                Document doc = uidoc.Document;

                List<AROpeningHoleEntity> arOHEColl = SelectAndGetAROpenHoles(uidoc, doc);
                if (arOHEColl == null)
                    return;

                AROpeningHoleEntity unitAR_OHE = CreateUnitOpeningHole(doc, arOHEColl);

                if (unitAR_OHE != null)
                {
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new AR_OHE_ElementDeleter(arOHEColl));
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new AR_OHE_Maker(new List<AROpeningHoleEntity>() { unitAR_OHE }));
                }
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
        /// Получить коллекцию отверстий в модели АР
        /// </summary>
        private List<AROpeningHoleEntity> SelectAndGetAROpenHoles(UIDocument uidoc, Document doc)
        {
            List<AROpeningHoleEntity> arOHEColl = new List<AROpeningHoleEntity>();

            MechEquipDocSelectionFilter selectionFilter = new MechEquipDocSelectionFilter();

            IList<Reference> pickedRefs;
            try
            {
                // Запуск выбару элементаў карыстальнікам
                pickedRefs = uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    selectionFilter,
                    "Выделите отверстия АР, которые нужно объеденить в одно");
            }
            // Отмена пользователем
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return null; }

            foreach (Reference reference in pickedRefs)
            {
                Element selectedElem = doc.GetElement(reference.ElementId);
                if (selectedElem == null)
                    continue;

                Element hostElem = null;
                if (!(selectedElem is FamilyInstance fi))
                {
                    HtmlOutput.Print(
                        $"Выбранное отверстия с id: {selectedElem.Id} не это не семейство. Обратись к разработчику, и выполни объединение вручную",
                        MessageType.Error);
                    return null;
                }
                else
                    hostElem = fi.Host;

                if (hostElem == null)
                {
                    HtmlOutput.Print(
                        $"Для отверстия с id: {selectedElem.Id} не удалось найти основу (стену/каркас несущий). Выполни объединение вручную",
                        MessageType.Error);
                    return null;
                }

                if (!(selectedElem.Location is LocationPoint locPnt))
                {
                    HtmlOutput.Print(
                        $"Для отверстия с id: {selectedElem.Id} не удалось найти точку вставки. Выполни объединение вручную",
                        MessageType.Error);
                    return null;
                }

                AROpeningHoleEntity arOHE = new AROpeningHoleEntity(Core.MainEntity.OpenigHoleShape.Rectangle, fi.Symbol.Name, hostElem, locPnt.Point, selectedElem);

                arOHE.SetFamilyPathAndName(doc);
                arOHE.SetGeomParams();
                arOHEColl.Add(arOHE);
            }

            return arOHEColl;
        }

        /// <summary>
        /// Создать объединённое отверстие по выбранной коллекции
        /// </summary>
        private AROpeningHoleEntity CreateUnitOpeningHole(Document doc, List<AROpeningHoleEntity> arOHEColl)
        {
            // Подбираю результирующий тип для семейства
            string resultSubDep = string.Empty;
            IEnumerable<string> arOHESubDeps = arOHEColl.Select(ohe => ohe.AR_OHE_IOSDubDepCode);
            if (arOHESubDeps.Distinct().Count() > 1 || arOHESubDeps.All(subDep => subDep.Equals("Несколько категорий")))
                resultSubDep = "Несколько категорий";
            else
                resultSubDep = arOHESubDeps.FirstOrDefault();

            
            // Анализирую на наличие нескольких основ у выборки отверстий
            IEnumerable<int> arOHEHostId = arOHEColl.Select(ohe => ohe.AR_OHE_HostElement.Id.IntegerValue);
            if (arOHEHostId.Distinct().Count() > 1)
            {
                HtmlOutput.Print(
                    $"Выбранные отверстия относятся к разным основаниям. Можно объединять отверстия ТОЛЬКО в рамках одной стены. Проанализируй корректность, и выполни объединение вручную",
                    MessageType.Error);
                return null;
            }

            // Получаю единственную основу
            Element hostElem = doc.GetElement(new ElementId(arOHEHostId.FirstOrDefault()));

            // Анализирую сущности и нахожу результирующий размер
            Solid unionSolid = null;
            foreach (AROpeningHoleEntity arOHE in arOHEColl)
            {
                try
                {
                    if (unionSolid == null)
                        unionSolid = arOHE.OHE_Solid;
                    else
                    {
                        Solid tempUnionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(unionSolid, arOHE.OHE_Solid, BooleanOperationsType.Union);
                        if (tempUnionSolid != null && tempUnionSolid.Volume > 0)
                            unionSolid = tempUnionSolid;
                        
                    }
                }
                catch (Exception ex)
                {
                    // Актуально для семейств, у которых нет тела (т.е. просто отверстия, например для СЕТ)
                    if (ex.Message.Contains("проблемы с получением Solid. Отправь разработчику"))
                    {

                    }
                    else
                        throw;
                }
            }

            // Анализирую вектор основы
            XYZ hostDir = GetHostDirection(hostElem);

            // Получаю координаты экстремумов
            XYZ[] minMaxPnts = GetMinAndMaxSolidPoints(unionSolid, hostDir);





            // Получаю высоту/шиирину
            Face hostFace = GetHostFace(hostElem, hostDir);
            Surface hostSurface = hostFace.GetSurface();
            
            hostSurface.Project(minMaxPnts[0], out UV minUV, out double minDist);
            hostSurface.Project(minMaxPnts[1], out UV maxUV, out double maxDist);
            if (minUV == null || maxUV == null)
                throw new Exception("Не удалось осуществить проекцию на стену. Отправь разработчику!");

            double resultWidth = maxUV.U - minUV.U;
            double resultHeight = maxUV.V - minUV.V;

            // Создаю точку вставки
            XYZ unionSolidCentroid = unionSolid.ComputeCentroid();
            BoundingBoxXYZ unionSolidBBox = unionSolid.GetBoundingBox();
            Transform bboxTrans = unionSolidBBox.Transform;

            double locPointX = (bboxTrans.OfPoint(unionSolidBBox.Min).X + bboxTrans.OfPoint(unionSolidBBox.Max).X) / 2;
            double locPointY = (bboxTrans.OfPoint(unionSolidBBox.Min).Y + bboxTrans.OfPoint(unionSolidBBox.Max).Y) / 2;
            double locPointZ = minMaxPnts[0].Z;
            XYZ locPoint = new XYZ(locPointX, locPointY, locPointZ);

            // Создаю сущность для заполнения
            AROpeningHoleEntity result = new AROpeningHoleEntity(
                Core.MainEntity.OpenigHoleShape.Rectangle,
                resultSubDep,
                hostElem,
                locPoint);

            result.SetFamilyPathAndName(doc);
            result.SetGeomParams();
            result.SetGeomParamsRoundData(resultHeight, resultWidth, 0);

            return result;
        }

        /// <summary>
        /// Получить вектор для основы отверстия
        /// </summary>
        private XYZ GetHostDirection(Element host)
        {
            // Получаю вектор для стены
            XYZ wallDirection;
            if (host is Wall wall)
            {
                Curve curve = (wall.Location as LocationCurve).Curve ??
                    throw new Exception($"Не обработанная основа (не Curve) с id: {host.Id}. Отправь разработчику!");

                XYZ wallOrigin = curve.GetEndPoint(0);
                XYZ wallEndPoint = curve.GetEndPoint(1);
                wallDirection = wallEndPoint - wallOrigin;
            }
            else
                throw new Exception($"Не обработанная основа с id: {host.Id}. Отправь разработчику!");


            return wallDirection;
        }


        /// <summary>
        /// Получить вектор для основы отверстия
        /// </summary>
        private Face GetHostFace(Element host, XYZ hostDirection)
        {
            Solid hostSolid = GeometryWorker.GetElemSolid(host);
            foreach (Face face in hostSolid.Faces)
            {
                if (face is PlanarFace planarFace)
                {
                    XYZ faceNormal = planarFace.FaceNormal;
                    XYZ checkOrigin = new XYZ(faceNormal.X, faceNormal.Y, hostDirection.Z);

                    double angle = UnitUtils.ConvertFromInternalUnits(hostDirection.AngleTo(checkOrigin), DisplayUnitType.DUT_DEGREES_AND_MINUTES);
                    if (Math.Round(angle, 5) == 90)
                        return face;
                }
            }

            return null;
        }


        /// <summary>
        /// Получить массив из МИНИМАЛЬНО и МАКСИМАЛЬНОЙ (именно в таком порядке) точек для солида с проекцией на вектор стены
        /// </summary>
        private XYZ[] GetMinAndMaxSolidPoints(Solid inputSolid, XYZ hostDirection)
        {
            double minX, minY, minZ;
            double maxX, maxY, maxZ;
            minX = minY = minZ =  double.MaxValue;
            maxX = maxY = maxZ = double.MinValue;

            // Получаю параллельную поверхность
            IList<Solid> splitedSolids = SolidUtils.SplitVolumes(inputSolid);
            foreach (Solid solid in splitedSolids)
            {
                Face checkFace = null;
                foreach (Face face in solid.Faces)
                {
                    if (face is PlanarFace planarFace)
                    {
                        XYZ faceNormal = planarFace.FaceNormal;
                        XYZ checkOrigin = new XYZ(faceNormal.X, faceNormal.Y, hostDirection.Z);

                        double angle = UnitUtils.ConvertFromInternalUnits(hostDirection.AngleTo(checkOrigin), DisplayUnitType.DUT_DEGREES_AND_MINUTES);
                        if (Math.Round(angle, 5) == 90)
                            checkFace = face;
                    }
                }

                if (checkFace == null)
                    throw new Exception($"Не удалось подобрать необходимую плоскость для отверстия. Отправь разработчику!");

                // Триангулирую поверхность и нахажу мин и макс точки
                Mesh chFaceMesh = checkFace.Triangulate(1);
                IList<XYZ> meshVertices = chFaceMesh.Vertices;
                foreach (XYZ vertex in meshVertices)
                {
                    if (vertex.X > maxX) maxX = vertex.X;
                    if (vertex.Y > maxY) maxY = vertex.Y;
                    if (vertex.Z > maxZ) maxZ = vertex.Z;

                    if (vertex.X < minX) minX = vertex.X;
                    if (vertex.Y < minY) minY = vertex.Y;
                    if (vertex.Z < minZ) minZ = vertex.Z;
                }
            }

            // Проецирую полученные минимумы и максимумы с трансофрмом на вектор
            Transform wallTrans = Transform.CreateTranslation(hostDirection);

            return new XYZ[2] 
            {
                wallTrans.OfPoint(new XYZ(minX, minY, minZ)),
                wallTrans.OfPoint(new XYZ(maxX, maxY, maxZ))
            };
        }

        /// <summary>
        /// Получить основную поверхность, которая будет параллельна хотсу (лицевая часть)
        /// </summary>
        /// <param name="solid"></param>
        /// <param name="host"></param>
        /// <returns></returns>
        private Face GetOHEMainFaceByHost(Solid solid, Element host)
        {
            // Получаю вектор для стены
            XYZ wallDirection;
            if (host is Wall wall)
            {
                Curve curve = (wall.Location as LocationCurve).Curve 
                    ?? throw new Exception($"Не обработанная основа (не Curve) с id: {host.Id}. Отправь разработчику!");
                
                XYZ wallOrigin = curve.GetEndPoint(0);
                XYZ wallEndPoint = curve.GetEndPoint(1);
                wallDirection = wallEndPoint - wallOrigin;
            }
            else
                throw new Exception($"Не обработанная основа с id: {host.Id}. Отправь разработчику!");


            foreach (Face face in solid.Faces)
            {
                PlanarFace planarFace = face as PlanarFace;
                
                XYZ faceNormal = planarFace.FaceNormal;
                XYZ checkOrigin = new XYZ(faceNormal.X, faceNormal.Y, wallDirection.Z);
                
                double angle = UnitUtils.ConvertFromInternalUnits(wallDirection.AngleTo(checkOrigin), DisplayUnitType.DUT_DEGREES_AND_MINUTES);
                if (Math.Round(angle, 5) == 90)
                    return face;
            }

            return null;
        }

        #endregion

        private void OnPropertyChanged([CallerMemberName] string name = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
