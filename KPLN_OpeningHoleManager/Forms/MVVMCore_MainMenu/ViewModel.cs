using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_OpeningHoleManager.Core;
using KPLN_OpeningHoleManager.Core.MainEntity;
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
    public class ViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private double _openHoleExpandedValue = 50;

        private double _ar_openHoleMinWidthValue = 300;
        private double _ar_openHoleMinHeightValue = 300;
        private double _ar_openHoleMinDistanceValue = 200;

        private double _kr_openHoleMinWidthValue = 200;
        private double _kr_openHoleMinHeightValue = 200;
        private double _kr_openHoleMinDistanceValue = 100;

        /// <summary>
        /// Комманда: Расставить отверстия по элементам ИОС
        /// </summary>
        public ICommand CreateOpenHoleByIOSElemsCommand { get; }

        /// <summary>
        /// Комманда: Объеденить отверстия
        /// </summary>
        public ICommand UnitOpenHolesCommand { get; }

        /// <summary>
        /// Комманда: Выбрать задания и расстваить отверстия в модели
        /// </summary>
        public ICommand SetOpenHoleByTaskCommand { get; }

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

        /// <summary>
        /// АР: Значение минимальной ширины отверстия при расстановке
        /// </summary>
        public double AR_OpenHoleMinWidthValue
        {
            get => _ar_openHoleMinWidthValue;
            set
            {
                _ar_openHoleMinWidthValue = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// АР: Значение минимальной высоты отверстия при расстановке
        /// </summary>
        public double AR_OpenHoleMinHeightValue
        {
            get => _ar_openHoleMinHeightValue;
            set
            {
                _ar_openHoleMinHeightValue = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// АР: Значение минимального расстояни между отверстиями для объединения (UV)
        /// </summary>
        public double AR_OpenHoleMinDistanceValue
        {
            get => _ar_openHoleMinDistanceValue;
            set
            {
                _ar_openHoleMinDistanceValue = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// КР: Значение минимальной ширины отверстия при расстановке
        /// </summary>
        public double KR_OpenHoleMinWidthValue
        {
            get => _kr_openHoleMinWidthValue;
            set
            {
                _kr_openHoleMinWidthValue = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// КР: Значение минимальной высоты отверстия при расстановке
        /// </summary>
        public double KR_OpenHoleMinHeightValue
        {
            get => _kr_openHoleMinHeightValue;
            set
            {
                _kr_openHoleMinHeightValue = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// КР: Значение минимального расстояни между отверстиями для объединения (UV)
        /// </summary>
        public double KR_OpenHoleMinDistanceValue
        {
            get => _kr_openHoleMinDistanceValue;
            set
            {
                _kr_openHoleMinDistanceValue = value;
                OnPropertyChanged();
            }
        }

        public ViewModel()
        {
            CreateOpenHoleByIOSElemsCommand = new RelayCommand(CreateOpenHoleByIOSElems);
            UnitOpenHolesCommand = new RelayCommand(UnitOpenHoles);
            SetOpenHoleByTaskCommand = new RelayCommand(SetOpenHoleByTask);
        }

        #region Расставить отверстия по элементам ИОС
        /// <summary>
        /// Реализация: Расставить отверстия по элементам ИОС
        /// </summary>
        private void CreateOpenHoleByIOSElems()
        {
            try
            {
                if (Module.CurrentUIApplication == null) return;

                UIDocument uidoc = Module.CurrentUIApplication.ActiveUIDocument;
                Document doc = uidoc.Document;


                Element selectedHost = SelectARHost(uidoc, doc);
                if (selectedHost == null)
                    return;


                IOSElemEntity[] iosEntities = GetIOSElemsFromLink(doc, selectedHost);
                if (iosEntities.Count() == 0)
                {
                    new TaskDialog("Внимание")
                    {
                        MainIcon = TaskDialogIcon.TaskDialogIconWarning,
                        MainInstruction = $"У выбранного элемента нет коллизий с ИОС, для создания отверстий",
                    }.Show();

                    return;
                }

                AROpeningHoleEntity[] arEntities = GetAROpeningsFromIOSElems_WithoutUnion(doc, selectedHost, iosEntities);

                AROpeningHoleEntity[] arEntitiesForUnion = GetAROpeningsFromIOSElems_UnionByParams(arEntities);
                if (arEntitiesForUnion.Any())
                {
                    List<AROpeningHoleEntity> tempAREntities = new List<AROpeningHoleEntity>(arEntitiesForUnion.Count());
                    foreach(AROpeningHoleEntity arEnt_Union in arEntitiesForUnion)
                    {
                        tempAREntities.Add(CreateUnitOpeningHole(doc, arEntitiesForUnion));
                    }

                    AROpeningHoleEntity[] resultAREntities = tempAREntities.ToArray();
                    if (resultAREntities != null)
                    {
                        KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new AR_OHE_Maker(ref resultAREntities));
                        KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new AR_OHE_ElementDeleter(resultAREntities));
                    }
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
        /// Выбрать основание для отверстия
        /// </summary>
        private Element SelectARHost(UIDocument uidoc, Document doc)
        {
            ARHostDocSelectionFilter selectionFilter = new ARHostDocSelectionFilter();

            Reference pickedRefs = null;
            try
            {
                // Запуск выбару элементаў карыстальнікам
                pickedRefs = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    selectionFilter,
                    "Выберите основание, в котором нужно добавить отверстия АР");
            }
            // Отмена пользователем
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return null; }
            // Отмена пользователем
            // Отмена пользователем
            catch (Autodesk.Revit.Exceptions.InvalidOperationException ioe)
            {
                if (ioe.Message.Contains("Cannot re-enter the pick operation"))
                {
                    new TaskDialog("Ошибка")
                    {
                        MainIcon = TaskDialogIcon.TaskDialogIconError,
                        MainInstruction = $"Сначала заверши предыдущий выбор, прежде чем начинать новый",
                    }.Show();
                    return null;
                }
            }

            return doc.GetElement(pickedRefs.ElementId);
        }

        /// <summary>
        /// Получить коллекцию эл-в ИОС для анализа по выбранному основанию
        /// </summary>
        private IOSElemEntity[] GetIOSElemsFromLink(Document doc, Element selectedHostElem)
        {
            List<IOSElemEntity> iosEntities = new List<IOSElemEntity>();

            Solid selectedHostElemSolid = GeometryWorker.GetElemSolid(selectedHostElem);

            // Анализирую ИОС файлы, и забираю из них элементы для анализа на коллизии
            RevitLinkInstance[] revitLinkInsts = new FilteredElementCollector(doc)
               .OfClass(typeof(RevitLinkInstance))
               .WhereElementIsNotElementType()
               .Cast<RevitLinkInstance>()
               .ToArray();

            DocumentSet docSet = doc.Application.Documents;
            foreach (Document openDoc in docSet)
            {
                if (!openDoc.IsLinked || openDoc.Title == doc.Title)
                    continue;
                try
                {
                    // Проверяю по типу отдела файла, стоит ли его анализировать
                    int openDocPrjDBSubDepartmentId = MainDBService.Get_DBDocumentSubDepartment(openDoc).Id;
                    switch (openDocPrjDBSubDepartmentId)
                    {
                        case 2:
                            goto case 3;
                        case 3:
                            break;
                        default:
                            RevitLinkInstance rLink = revitLinkInsts
                                .FirstOrDefault(rl => openDoc.Title.Contains(rl.Name.Split(new string[] { ".rvt" }, StringSplitOptions.None)
                                .FirstOrDefault()));

                            // Если открыто сразу несколько моделей одного проекта, то линки могут прилететь с другого файла. В таком случае - игнор
                            if (rLink != null)
                                iosEntities.AddRange(IOSElemsCollectionCreator.CreateIOSElemEntitiesForLink_BySolid(rLink, selectedHostElemSolid));

                            break;

                    }
                }
                catch (Exception ex)
                {
                    HtmlOutput.Print($"Ошибка: {ex.Message}", MessageType.Error);
                }
            }

            return iosEntities.ToArray();
        }

        /// <summary>
        /// Получить коллекцию отверстий АР по элементам ИОС БЕЗ объединения
        /// </summary>
        /// <returns></returns>
        private AROpeningHoleEntity[] GetAROpeningsFromIOSElems_WithoutUnion(Document doc, Element hostElem, IOSElemEntity[] iosElemColl)
        {
            List<AROpeningHoleEntity> result = new List<AROpeningHoleEntity>();

            XYZ hostDir = GetHostDirection(hostElem);
            Face hostDirFace = GetFace_ParalleledForDirection(GeometryWorker.GetElemSolid(hostElem), hostDir);
            Surface hostSurface = hostDirFace.GetSurface();

            // Одиночная обработка каждого эл-та
            foreach (IOSElemEntity iosElemEnt in iosElemColl)
            {
                // Получаю форму одиночного отверстия
                OpenigHoleShape ohe_Shape = OpenigHoleShape.Circle;
                ConnectorSet conSet = null;
                if (iosElemEnt.IOS_Element is MEPCurve mc)
                    conSet = mc.ConnectorManager.Connectors;
                else if (iosElemEnt.IOS_Element is FamilyInstance fi && fi.MEPModel is MechanicalFitting mf)
                    conSet = mf.ConnectorManager.Connectors;

                if (conSet != null)
                {
                    foreach (Connector con in conSet)
                    {
                        if (con.Shape == ConnectorProfileType.Rectangular)
                        {
                            ohe_Shape = OpenigHoleShape.Rectangle;
                            break;
                        }
                    }
                }


                // Получаю координаты экстремумов
                XYZ[] minMaxPnts = GetMinAndMaxSolidPoints(iosElemEnt.ARIOS_IntesectionSolid, hostDir);


                // Получаю координаты центра пересечения
                XYZ docCoord = iosElemEnt.IOS_LinkTransform.OfPoint(iosElemEnt.ARIOS_IntesectionSolid.ComputeCentroid());


                // Получаю высоту/ширину пересечения
                hostSurface.Project(minMaxPnts[0], out UV minUV, out double minDist);
                hostSurface.Project(minMaxPnts[1], out UV maxUV, out double maxDist);
                if (minUV == null || maxUV == null)
                    throw new Exception("Не удалось осуществить проекцию на отверстие. Отправь разработчику!");

                double resultWidth = Math.Abs(maxUV.U - minUV.U);
                double resultHeight = Math.Abs(maxUV.V - minUV.V);
                double resultRadius = Math.Abs(maxUV.V - minUV.V);


                // Создаю сущность AROpeningHoleEntity
                AROpeningHoleEntity arEntity = new AROpeningHoleEntity(ohe_Shape, MainDBService.Get_DBDocumentSubDepartment(iosElemEnt.IOS_LinkDocument).Code, hostElem, docCoord)
                    .SetFamilyPathAndName(doc)
                    as AROpeningHoleEntity;

                arEntity.SetGeomParams();
                arEntity.SetGeomParamsRoundData(resultWidth, resultHeight, resultRadius, OpenHoleExpandedValue);
                arEntity.UpdatePointData_ByShape();


                // Фильтрую отверстия в одной точке c меньшими размерами - они идут в игнор, т.к. ничего не вырежут
                if (result.Any(arEnt => arEnt.OHE_Point.IsAlmostEqualTo(arEntity.OHE_Point, 0.01)
                    && arEnt.OHE_Height <= arEntity.OHE_Height
                    && arEnt.OHE_Width <= arEntity.OHE_Width
                    && arEnt.OHE_Radius <= arEntity.OHE_Radius))
                    continue;
                else
                    result.Add(arEntity);
            }

            return result.ToArray();
        }

        /// <summary>
        /// Получить коллекцию отверстий АР, которые могут быть объеденены по нужным параметрам
        /// </summary>
        /// <returns></returns>
        private AROpeningHoleEntity[] GetAROpeningsFromIOSElems_UnionByParams(AROpeningHoleEntity[] arEntities)
        {
            List<AROpeningHoleEntity> result = new List<AROpeningHoleEntity>();
            foreach (AROpeningHoleEntity arCheckIntersectEntity1 in arEntities)
            {
                List<AROpeningHoleEntity> intersectedAREntities = new List<AROpeningHoleEntity>();
                List<AROpeningHoleEntity> closeInDistAREntities = new List<AROpeningHoleEntity>();

                // Параметры для нахождения пересечений
                double volume1 = arCheckIntersectEntity1.OHE_Solid.Volume;
                XYZ centr1 = arCheckIntersectEntity1.OHE_Solid.ComputeCentroid();

                // Параметры для нахождения расстояния
                XYZ hostDir1 = GetHostDirection(arCheckIntersectEntity1.AR_OHE_HostElement);
                Face hostDirFace1 = GetFace_ParalleledForDirection(arCheckIntersectEntity1.OHE_Solid, hostDir1);
                Surface hostSurface1 = hostDirFace1.GetSurface();

                foreach (AROpeningHoleEntity arCheckIntersectEntity2 in arEntities)
                {
                    // Игнорирую эквивалентный солид
                    double volume2 = arCheckIntersectEntity2.OHE_Solid.Volume;
                    XYZ centr2 = arCheckIntersectEntity2.OHE_Solid.ComputeCentroid();
                    if (Math.Abs(volume1 - volume2) <= 0.01 && centr1.IsAlmostEqualTo(centr2, 0.01))
                        continue;


                    // Нахожу пересечение солидов
                    Solid intersectSolid = BooleanOperationsUtils.ExecuteBooleanOperation(arCheckIntersectEntity1.OHE_Solid, arCheckIntersectEntity2.OHE_Solid, BooleanOperationsType.Intersect);
                    if (intersectSolid != null && intersectSolid.Volume > 0)
                    {
                        intersectedAREntities.Add(arCheckIntersectEntity2);
                        break;
                    }


                    // Нахожу близкие по расстоянию элементы
                    var a = AR_OpenHoleMinDistanceValue;
                    var b = KR_OpenHoleMinDistanceValue;


                }


            }


            return result.ToArray();
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

                AROpeningHoleEntity[] arOHEColl = SelectAndGetAROpenHoles(uidoc, doc);
                if (arOHEColl == null)
                    return;

                AROpeningHoleEntity[] unionAR_OHE = new AROpeningHoleEntity[] { CreateUnitOpeningHole(doc, arOHEColl) };

                if (unionAR_OHE != null)
                {
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new AR_OHE_Maker(ref unionAR_OHE));
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new AR_OHE_ElementDeleter(unionAR_OHE));
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
        private AROpeningHoleEntity[] SelectAndGetAROpenHoles(UIDocument uidoc, Document doc)
        {
            List<AROpeningHoleEntity> arOHEColl = new List<AROpeningHoleEntity>();

            MechEquipDocSelectionFilter selectionFilter = new MechEquipDocSelectionFilter();

            IList<Reference> pickedRefs = null;
            try
            {
                // Запуск выбару элементаў карыстальнікам
                pickedRefs = uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    selectionFilter,
                    "Выберите отверстия АР, которые нужно объеденить в одно");
            }
            // Отмена пользователем
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return null; }
            // Отмена пользователем
            catch (Autodesk.Revit.Exceptions.InvalidOperationException ioe)
            {
                if (ioe.Message.Contains("Cannot re-enter the pick operation"))
                {
                    new TaskDialog("Ошибка")
                    {
                        MainIcon = TaskDialogIcon.TaskDialogIconError,
                        MainInstruction = $"Сначала заверши предыдущий выбор, прежде чем начинать новый",
                    }.Show();
                    return null;
                }
            }

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

            return arOHEColl.ToArray();
        }

        /// <summary>
        /// Создать объединённое отверстие по выбранной коллекции
        /// </summary>
        private AROpeningHoleEntity CreateUnitOpeningHole(Document doc, AROpeningHoleEntity[] arOHEColl)
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
            Solid hostSolid = GeometryWorker.GetElemSolid(hostElem);
            Face hostFace = GetFace_ParalleledForDirection(hostSolid, hostDir);
            Surface hostSurface = hostFace.GetSurface();

            hostSurface.Project(minMaxPnts[0], out UV minUV, out double minDist);
            hostSurface.Project(minMaxPnts[1], out UV maxUV, out double maxDist);
            if (minUV == null || maxUV == null)
                throw new Exception("Не удалось осуществить проекцию на стену. Отправь разработчику!");

            double resultWidth = Math.Abs(maxUV.U - minUV.U);
            double resultHeight = Math.Abs(maxUV.V - minUV.V);

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
                OpenigHoleShape.Rectangle,
                resultSubDep,
                hostElem,
                locPoint);

            result.SetFamilyPathAndName(doc);
            result.SetGeomParams();
            result.SetGeomParamsRoundData(resultHeight, resultWidth, 0);

            return result;
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

                IOSOpeningHoleTaskEntity[] iosTasks = SelectAndGetIOSTasksFromLink(uidoc, doc);
                if (!iosTasks.Any())
                    return;

                AROpeningHoleEntity[] arEntities = GetAROpeningsFromIOSTask(doc, iosTasks);

                if (arEntities.Any())
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new AR_OHE_Maker(ref arEntities));
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
        private IOSOpeningHoleTaskEntity[] SelectAndGetIOSTasksFromLink(UIDocument uidoc, Document doc)
        {
            List<IOSOpeningHoleTaskEntity> iosTasks = new List<IOSOpeningHoleTaskEntity>();

            MechEquipLinkSelectionFilter selectionFilter = new MechEquipLinkSelectionFilter(doc);

            IList<Reference> pickedRefs = null;
            try
            {
                // Запуск выбару элементаў карыстальнікам
                pickedRefs = uidoc.Selection.PickObjects(
                    ObjectType.LinkedElement,
                    selectionFilter,
                    "Выберите задания от ИОС, которые нужно превартить в отверстия АР");
            }
            // Отмена пользователем
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return null; }
            // Отмена пользователем
            // Отмена пользователем
            catch (Autodesk.Revit.Exceptions.InvalidOperationException ioe)
            {
                if (ioe.Message.Contains("Cannot re-enter the pick operation"))
                {
                    new TaskDialog("Ошибка")
                    {
                        MainIcon = TaskDialogIcon.TaskDialogIconError,
                        MainInstruction = $"Сначала заверши предыдущий выбор, прежде чем начинать новый",
                    }.Show();
                    return null;
                }
            }

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

            return iosTasks.ToArray();
        }

        /// <summary>
        /// Получить коллекцию отверстий АР из заданий от ИОС
        /// </summary>
        private AROpeningHoleEntity[] GetAROpeningsFromIOSTask(Document doc, IOSOpeningHoleTaskEntity[] iosTasks)
        {
            List<AROpeningHoleEntity> arEntities = new List<AROpeningHoleEntity>();

            Element[] arPotentialHosts = GetPotentialARHosts(doc, iosTasks);

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

                arEntity.SetGeomParams();
                arEntity.SetGeomParamsRoundData(iosTask.OHE_Height, iosTask.OHE_Width, iosTask.OHE_Radius);
                arEntity.UpdatePointData_ByShape();

                arEntities.Add(arEntity);
            }

            return arEntities.ToArray();
        }

        /// <summary>
        /// Получить список потенциальных основ для отверстия на основе выбранных заданий от ИОС
        /// </summary>
        private Element[] GetPotentialARHosts(Document doc, IOSOpeningHoleTaskEntity[] iosTasks)
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

            return result.ToArray();
        }

        /// <summary>
        /// Получить стену, в которую будем ставить отверстие
        /// </summary>
        /// <returns></returns>
        private Element GetHostForOpening(Element[] arPotentialHosts, IOSOpeningHoleTaskEntity iosTask)
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
        private Face GetFace_ParalleledForDirection(Solid checkSolid, XYZ hostDirection)
        {
            foreach (Face face in checkSolid.Faces)
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
            minX = minY = minZ = double.MaxValue;
            maxX = maxY = maxZ = double.MinValue;

            // Получаю параллельную поверхность
            IList<Solid> splitedSolids = SolidUtils.SplitVolumes(inputSolid);
            foreach (Solid solid in splitedSolids)
            {
                double tempFaceArea = 0;
                Face checkFace = null;
                foreach (Face face in solid.Faces)
                {
                    if (face is PlanarFace planarFace)
                    {
                        XYZ faceNormal = planarFace.FaceNormal;
                        XYZ checkOrigin = new XYZ(faceNormal.X, faceNormal.Y, hostDirection.Z);

                        double angle = UnitUtils.ConvertFromInternalUnits(hostDirection.AngleTo(checkOrigin), DisplayUnitType.DUT_DEGREES_AND_MINUTES);
                        if (Math.Round(angle, 5) == 90 && face.Area > tempFaceArea)
                        {
                            tempFaceArea = face.Area;
                            checkFace = face;
                        }
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

        private void OnPropertyChanged([CallerMemberName] string name = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
