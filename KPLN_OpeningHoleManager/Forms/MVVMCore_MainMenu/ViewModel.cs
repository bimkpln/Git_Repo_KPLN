using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
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
    public sealed class ViewModel : INotifyPropertyChanged
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
        /// Комманда: Расставить отверстия по ВСЕМ элементам ИОС
        /// </summary>
        public ICommand CreateOpenHoleByIOSElemsCommand { get; }

        /// <summary>
        /// Комманда: Расставить отверстия по ВЫБРАННЫМ элементам ИОС
        /// </summary>
        public ICommand CreateOpenHoleBySelectedIOSElemsCommand { get; }

        /// <summary>
        /// Комманда: Объеденить отверстия
        /// </summary>
        public ICommand UnionOpenHolesCommand { get; }

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
            // Устанавливаю команды
            CreateOpenHoleByIOSElemsCommand = new RelayCommand(CreateOpenHoleByIOSElems);
            CreateOpenHoleBySelectedIOSElemsCommand = new RelayCommand(CreateOpenHoleBySelectedIOSElems);
            UnionOpenHolesCommand = new RelayCommand(UnionOpenHoles);
            SetOpenHoleByTaskCommand = new RelayCommand(SetOpenHoleByTask);
        }

        /// <summary>
        /// Чтение и установка переменных (длина, ширина и т.п.) из файла-конфигурации (если он есть)
        /// </summary>
        public void SetDataGeomParamsData(DBProject dBProject)
        {
            MainConfig mainConfig = MainConfig.GetData_FromConfig(dBProject);
            if (mainConfig != null)
            {
                OpenHoleExpandedValue = mainConfig.OpenHoleExpandedValue;

                AR_OpenHoleMinDistanceValue = mainConfig.AR_OpenHoleMinDistanceValue;
                AR_OpenHoleMinHeightValue = mainConfig.AR_OpenHoleMinHeightValue;
                AR_OpenHoleMinWidthValue = mainConfig.AR_OpenHoleMinWidthValue;

                KR_OpenHoleMinDistanceValue = mainConfig.KR_OpenHoleMinDistanceValue;
                KR_OpenHoleMinHeightValue = mainConfig.KR_OpenHoleMinHeightValue;
                KR_OpenHoleMinWidthValue = mainConfig.KR_OpenHoleMinWidthValue;
            }
        }

        #region Расставить отверстия по ВСЕМ элементам ИОС
        /// <summary>
        /// Реализация: Расставить отверстия по ВСЕМ элементам ИОС
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

                AROpeningHoleEntity[] elemToCreate = GetAROpeningsFromIOSElems_WithoutUnion(doc, selectedHost, iosEntities);
                if (elemToCreate.Any())
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new AR_OHE_MakerWithUnion(elemToCreate, "KPLN: Отверстия по пересечению", this, false));
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
        /// Реализация: Расставить отверстия по ВЫБРАННЫМ элементам ИОС
        /// </summary>
        private void CreateOpenHoleBySelectedIOSElems()
        {
            try
            {
                if (Module.CurrentUIApplication == null) return;

                UIDocument uidoc = Module.CurrentUIApplication.ActiveUIDocument;
                Document doc = uidoc.Document;


                Element selectedHost = SelectARHost(uidoc, doc);
                if (selectedHost == null)
                    return;


                IOSElemEntity[] iosEntities = SelectAndGetIOSElemsFromLink(uidoc, doc, selectedHost);
                if (iosEntities == null)
                    return;

                if (iosEntities.Count() == 0)
                {
                    new TaskDialog("Внимание")
                    {
                        MainIcon = TaskDialogIcon.TaskDialogIconWarning,
                        MainInstruction = $"У выбранного элемента АР нет коллизий с выбранными элементами ИОС. Создать отверстия невозможно",
                    }.Show();

                    return;
                }

                AROpeningHoleEntity[] elemToCreate = GetAROpeningsFromIOSElems_WithoutUnion(doc, selectedHost, iosEntities);
                if (elemToCreate.Any())
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new AR_OHE_MakerWithUnion(elemToCreate, "KPLN: Отверстия по пересечению", this, false));
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
        /// Выбрать эл-ты ИОС и получить коллекцию эл-в для анализа по выбранному основанию
        /// </summary>
        private IOSElemEntity[] SelectAndGetIOSElemsFromLink(UIDocument uidoc, Document doc, Element selectedHostElem)
        {
            Solid selectedHostElemSolid = GeometryWorker.GetRevitElemSolid(selectedHostElem);

            List<IOSElemEntity> iosEntities = new List<IOSElemEntity>();

            IOSElemsLinkSelectionFilter selectionFilter = new IOSElemsLinkSelectionFilter(doc);

            IList<Reference> pickedRefs = null;
            try
            {
                // Запуск выбару элементаў карыстальнікам
                pickedRefs = uidoc.Selection.PickObjects(
                    ObjectType.LinkedElement,
                    selectionFilter,
                    "Выберите элементы ИОС, по которым нужно расставить отверстия в выбранной стене");
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


                IOSElemEntity iosElemEntity = IOSElemsCollectionCreator.CreateIOSElemEntityForLinkElem_BySolid(linkInstance, linkedElement, selectedHostElemSolid);
                if (iosElemEntity != null)
                    iosEntities.Add(iosElemEntity);
            }

            return iosEntities.ToArray();
        }

        /// <summary>
        /// Получить коллекцию эл-в ИОС для анализа по выбранному основанию
        /// </summary>
        private IOSElemEntity[] GetIOSElemsFromLink(Document doc, Element selectedHostElem)
        {
            List<IOSElemEntity> iosEntities = new List<IOSElemEntity>();

            Solid selectedHostElemSolid = GeometryWorker.GetRevitElemSolid(selectedHostElem);

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
        /// Получить коллекцию отверстий АР по элементам ИОС
        /// </summary>
        /// <returns></returns>
        private AROpeningHoleEntity[] GetAROpeningsFromIOSElems_WithoutUnion(Document doc, Element hostElem, IOSElemEntity[] iosElemColl)
        {
            List<AROpeningHoleEntity> result = new List<AROpeningHoleEntity>();

            XYZ hostDir = GeometryWorker.GetHostDirection(hostElem);

            // Одиночная обработка каждого эл-та
            foreach (IOSElemEntity iosElemEnt in iosElemColl)
            {
                // Получаю форму одиночного отверстия
                OpenigHoleShape ohe_Shape = OpenigHoleShape.Circle;
                ConnectorSet conSet = null;
                if (iosElemEnt.IOS_Element is MEPCurve mc)
                    conSet = mc.ConnectorManager.Connectors;
                else if (iosElemEnt.IOS_Element is FamilyInstance fi && fi.MEPModel is MEPModel mepMod)
                    conSet = mepMod.ConnectorManager.Connectors;

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


                // Получаю ширину и высоту
                double[] widthAndHeight = GeometryWorker.GetSolidWidhtAndHeight_ByDirection(iosElemEnt.ARIOS_IntesectionSolid, hostDir);
                double resultWidth = widthAndHeight[0];
                double resultHeight = widthAndHeight[1];
                double resultRadius = widthAndHeight[1];


                // Получаю координаты центра пересечения
                XYZ docCoord = iosElemEnt.ARIOS_IntesectionSolid.ComputeCentroid();


                // Создаю сущность AROpeningHoleEntity
                AROpeningHoleEntity arEntity = new AROpeningHoleEntity(ohe_Shape, MainDBService.Get_DBDocumentSubDepartment(iosElemEnt.IOS_LinkDocument).Code, hostElem, docCoord)
                    .SetFamilyPathAndName(doc)
                    as AROpeningHoleEntity;

                arEntity.SetGeomParams();
                arEntity.SetGeomParamsRoundData(resultHeight, resultWidth, resultRadius, OpenHoleExpandedValue);
                arEntity.UpdatePointData_ByShape();

                result.Add(arEntity);
            }

            return result.ToArray();
        }
        #endregion

        #region Объеденить отверстия
        /// <summary>
        /// Реализация: Объеденить отверстия
        /// </summary>
        private void UnionOpenHoles()
        {
            try
            {
                if (Module.CurrentUIApplication == null) return;

                UIDocument uidoc = Module.CurrentUIApplication.ActiveUIDocument;
                Document doc = uidoc.Document;


                AROpeningHoleEntity[] arOHEColl = SelectAndGetAROpenHoles(uidoc, doc);
                if (arOHEColl == null)
                    return;

                AROpeningHoleEntity[] unionAR_OHE = new AROpeningHoleEntity[] { AROpeningHoleEntity.CreateUnionOpeningHole(doc, arOHEColl) };
                if (unionAR_OHE != null)
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new AR_OHE_MakerWithUnion(unionAR_OHE, "KPLN: Объединить отверстия", this, true));
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
                if (iosTasks == null || !iosTasks.Any())
                    return;

                AROpeningHoleEntity[] arEntities = GetAROpeningsFromIOSTask(doc, iosTasks);

                if (arEntities.Any())
                {
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new AR_OHE_Maker(arEntities, "KPLN: Отверстия по заданию"));
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
                .Where(e => ARKRElemsWorker.ElemCatLogicalOrFilter.PassesFilter(e) && ARKRElemsWorker.ElemExtraFilterFunc(e)));

            result.AddRange(allHostFromDocumentColl
                .Where(e => ARKRElemsWorker.ElemCatLogicalOrFilter.PassesFilter(e) && ARKRElemsWorker.ElemExtraFilterFunc(e)));

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
                Solid arPotHostSolid = GeometryWorker.GetRevitElemSolid(arPotHost);
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

        private void OnPropertyChanged([CallerMemberName] string name = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
