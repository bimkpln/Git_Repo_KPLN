using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_OpeningHoleManager.Core;
using KPLN_OpeningHoleManager.Core.MainEntity;
using KPLN_OpeningHoleManager.ExecutableCommand;
using KPLN_OpeningHoleManager.Forms.MVVMCommand;
using KPLN_OpeningHoleManager.Forms.SelectionFilters;
using KPLN_OpeningHoleManager.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;

namespace KPLN_OpeningHoleManager.Forms.MVVMCore_MainMenu
{
    public sealed class MainViewModel : INotifyPropertyChanged
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
        /// Комманда: Расставить отверстия по ВСЕМ элементам ИОС на все элементы АР (кладка)
        /// </summary>
        public ICommand CreateOpenHole_AllARElemsCommand { get; }

        /// <summary>
        /// Комманда: Расставить отверстия по ВСЕМ элементам ИОС на выбранный эл-т АР/КР
        /// </summary>
        public ICommand CreateOpenHole_SelectedARKRAndIOSElemsCommand { get; }

        /// <summary>
        /// Комманда: Расставить отверстия по ВЫБРАННЫМ элементам ИОС на выбранный эл-т АР/КР
        /// </summary>
        public ICommand CreateOpenHole_SelectedARKRAndSelectedIOSElemsCommand { get; }

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

        public MainViewModel()
        {
            // Устанавливаю команды
            CreateOpenHole_AllARElemsCommand = new RelayCommand(CreateOpenHole_AllARElems);
            CreateOpenHole_SelectedARKRAndIOSElemsCommand = new RelayCommand(CreateOpenHole_SelectedARKRAndIOSElems);
            CreateOpenHole_SelectedARKRAndSelectedIOSElemsCommand = new RelayCommand(CreateOpenHole_SelectedARKRAndSelectedIOSElems);
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

        #region Расставить отверстия по ВСЕМ элементам ИОС на все элементы АР
        /// <summary>
        /// Реализация: Расставить отверстия по ВСЕМ элементам ИОС на все элементы АР (кладка)
        /// </summary>
        private void CreateOpenHole_AllARElems()
        {
            try
            {
                if (Module.CurrentUIApplication == null) return;

                Document doc = Module.CurrentUIApplication.ActiveUIDocument.Document;

                View activeView = Module.CurrentUIApplication.ActiveUIDocument.ActiveView;

                Element[] arHosts = new FilteredElementCollector(doc, activeView.Id)
                    .WhereElementIsNotElementType()
                    .Where(e => ARKRElemsWorker.ElemCatLogicalOrFilter.PassesFilter(e) && ARKRElemsWorker.ARElemExtraFilterFunc(e))
                    .ToArray();

                if (arHosts == null || CheckWSAvailableError(doc, arHosts.Select(el => el.Id)))
                    return;

                ProgressInfoViewModel progressInfoViewModel = new ProgressInfoViewModel();
                ProgressWindow window = new ProgressWindow(progressInfoViewModel);
                window.Show();

                ARKRElemEntity[] arkrEntities = GetARKRElemsFromLink(doc, arHosts, progressInfoViewModel);

                if (arkrEntities.Count() == 0)
                {
                    new TaskDialog("Внимание")
                    {
                        MainIcon = TaskDialogIcon.TaskDialogIconWarning,
                        MainInstruction = $"У элемента/-ов на виде нет коллизий с ИОС, для создания отверстий.\n" +
                            $"ВАЖНО: анализируются только перегородки АР.",
                    }.Show();

                    return;
                }

                AROpeningHoleEntity[] elemToCreate = AROpeningHoleEntity.ClearCollectionByJoinedHosts(doc, GetAROpeningsFromARKRElems_WithoutUnion(doc, arkrEntities, progressInfoViewModel));
                if (elemToCreate.Any())
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new AR_OHE_MakerWithUnion(elemToCreate, "KPLN: Отверстия по пересечению", this, false, progressInfoViewModel));
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
        #endregion

        #region Расставить отверстия по элементам ИОС на выбранную стену АР/КР
        /// <summary>
        /// Реализация: Расставить отверстия по ВСЕМ элементам ИОС
        /// </summary>
        private void CreateOpenHole_SelectedARKRAndIOSElems()
        {
            try
            {
                if (Module.CurrentUIApplication == null) return;

                UIDocument uidoc = Module.CurrentUIApplication.ActiveUIDocument;
                Document doc = uidoc.Document;

                Element[] selectedHosts = new Element[] { SelectARHost(uidoc, doc) };
                if (selectedHosts.FirstOrDefault() == null || CheckWSAvailableError(doc, selectedHosts.Select(el => el.Id)))
                    return;

                ProgressInfoViewModel progressInfoViewModel = new ProgressInfoViewModel();
                ProgressWindow window = new ProgressWindow(progressInfoViewModel);
                window.Show();

                ARKRElemEntity[] arkrEntities = GetARKRElemsFromLink(doc, selectedHosts, progressInfoViewModel);
                if (arkrEntities.Count() == 0)
                {
                    window.Close();

                    new TaskDialog("Внимание")
                    {
                        MainIcon = TaskDialogIcon.TaskDialogIconWarning,
                        MainInstruction = $"У выбранного элемента нет коллизий с ИОС, для создания отверстий",
                    }.Show();

                    return;
                }

                AROpeningHoleEntity[] elemToCreate = AROpeningHoleEntity.ClearCollectionByJoinedHosts(doc, GetAROpeningsFromARKRElems_WithoutUnion(doc, arkrEntities, progressInfoViewModel));
                if (elemToCreate.Any())
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new AR_OHE_MakerWithUnion(elemToCreate, "KPLN: Отверстия по пересечению", this, false, progressInfoViewModel));
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
        private void CreateOpenHole_SelectedARKRAndSelectedIOSElems()
        {
            try
            {
                if (Module.CurrentUIApplication == null) return;

                UIDocument uidoc = Module.CurrentUIApplication.ActiveUIDocument;
                Document doc = uidoc.Document;


                Element selectedHost = SelectARHost(uidoc, doc);
                if (selectedHost == null || CheckWSAvailableError(doc, new List<ElementId> { selectedHost.Id }))
                    return;

                ARKRElemEntity arkrEntity = SelectAndGetARKRElemsFromLink(uidoc, doc, selectedHost);
                if (arkrEntity == null)
                    return;

                ProgressInfoViewModel progressInfoViewModel = new ProgressInfoViewModel();
                ProgressWindow window = new ProgressWindow(progressInfoViewModel);
                window.Show();


                if (arkrEntity.IOSElemEntities.Count() == 0)
                {
                    window.Close();

                    new TaskDialog("Внимание")
                    {
                        MainIcon = TaskDialogIcon.TaskDialogIconWarning,
                        MainInstruction = $"У выбранного элемента АР нет коллизий с выбранными элементами ИОС. Создать отверстия невозможно",
                    }.Show();

                    return;
                }

                AROpeningHoleEntity[] elemToCreate = AROpeningHoleEntity.ClearCollectionByJoinedHosts(doc, GetAROpeningsFromARKRElems_WithoutUnion(doc, new ARKRElemEntity[] { arkrEntity }, progressInfoViewModel));
                if (elemToCreate.Any())
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new AR_OHE_MakerWithUnion(elemToCreate, "KPLN: Отверстия по пересечению", this, false, progressInfoViewModel));
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
                    "Выберите основание АР/КЖ, в котором нужно добавить отверстия");
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
        /// Проверка эл-в на занятый РН
        /// </summary>
        private bool CheckWSAvailableError(Document doc, IEnumerable<ElementId> elemIdsToCheck)
        {
            ICollection<ElementId> availableWSElemsId = WorksharingUtils.CheckoutElements(doc, elemIdsToCheck.ToArray());
            if (availableWSElemsId.Count < elemIdsToCheck.Count())
            {
                new TaskDialog("Ошибка")
                {
                    MainIcon = TaskDialogIcon.TaskDialogIconError,
                    MainInstruction = $"Возможность изменения ограничена для элементов. Попроси коллег ОСВОБОДИТЬ все забранные рабочие наборы и элементы\n",
                }.Show();

                return true;
            }

            return false;
        }

        /// <summary>
        /// Выбрать эл-ты ИОС и получить коллекцию эл-в для анализа по выбранному основанию
        /// </summary>
        private ARKRElemEntity SelectAndGetARKRElemsFromLink(UIDocument uidoc, Document doc, Element selectedHostElem)
        {
            ARKRElemEntity aRKRElemEntity = new ARKRElemEntity(selectedHostElem);

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

                ARKRElemsCollectionCreator.SetIOSEntities_BySelectedIOSElems(aRKRElemEntity, linkInstance, linkedElement);
            }

            return aRKRElemEntity;
        }

        /// <summary>
        /// Получить коллекцию эл-в ARRKEntities
        /// </summary>
        private ARKRElemEntity[] GetARKRElemsFromLink(Document doc, Element[] arHostElems, ProgressInfoViewModel progressInfoViewModel)
        {
            List<ARKRElemEntity> arkrEntities = new List<ARKRElemEntity>();

            // Анализирую ИОС файлы
            RevitLinkInstance[] revitLinkInsts = new FilteredElementCollector(doc)
               .OfClass(typeof(RevitLinkInstance))
               .WhereElementIsNotElementType()
               .Cast<RevitLinkInstance>()
               .ToArray();

            DocumentSet docSet = doc.Application.Documents;
            List<RevitLinkInstance> onlyIOSRLinkInsts = new List<RevitLinkInstance>();
            foreach (Document openDoc in docSet)
            {
                if (!openDoc.IsLinked || openDoc.Title == doc.Title)
                    continue;

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
                            onlyIOSRLinkInsts.Add(rLink);

                        break;
                }
            }

            // Генерирую коллекцию ARKRIOSElemEntity относительно основ АР/КР
            progressInfoViewModel.ProcessTitle = "Подготовка оснований...";
            progressInfoViewModel.MaxProgress = arHostElems.Length;
            foreach (Element arHostElem in arHostElems)
            {
                try
                {
                    ARKRElemEntity aRKRElemEntity = new ARKRElemEntity(arHostElem);

                    ARKRElemsCollectionCreator.SetIOSEntities_ByIOSElemEntities(aRKRElemEntity, onlyIOSRLinkInsts);
                    if (aRKRElemEntity.IOSElemEntities.Count() > 0)
                        arkrEntities.Add(aRKRElemEntity);
                }
                catch (Exception ex)
                {
                    HtmlOutput.Print($"Ошибка: {ex.Message}", MessageType.Error);
                }

                ++progressInfoViewModel.CurrentProgress;
                progressInfoViewModel.DoEvents();
            }

            return arkrEntities.ToArray();
        }

        /// <summary>
        /// Получить коллекцию отверстий АР по элементам ИОС
        /// </summary>
        /// <returns></returns>
        private AROpeningHoleEntity[] GetAROpeningsFromARKRElems_WithoutUnion(Document doc, ARKRElemEntity[] arkrElemColl, ProgressInfoViewModel progressInfoViewModel)
        {
            List<AROpeningHoleEntity> result = new List<AROpeningHoleEntity>();

            // Одиночная обработка каждого эл-та
            progressInfoViewModel.ProcessTitle = "Подготовка отверстий...";
            progressInfoViewModel.MaxProgress = arkrElemColl.Length;
            foreach (ARKRElemEntity arkrElemEnt in arkrElemColl)
            {
                XYZ hostDir = GeometryWorker.GetHostDirection(arkrElemEnt.ARKRHost_Element);

                foreach (IOSElemEntity iosElemEnt in arkrElemEnt.IOSElemEntities)
                {
                    // Получаю форму одиночного отверстия
                    OpenigHoleShape ohe_Shape = OpenigHoleShape.Rectangular;

                    Face intersectMainFace = GeometryWorker.GetFace_ByAngleToDirection(iosElemEnt.ARKRIOS_IntesectionSolid, hostDir)
                        // Такое может быть, если тело полность погружено в объём, тогда уменьшаем точность поиска
                        ?? GeometryWorker.GetFace_ByAngleToDirection(iosElemEnt.ARKRIOS_IntesectionSolid, hostDir, 5);

                    var edgeFIter = intersectMainFace.EdgeLoops.ForwardIterator();
                    bool moveIterator = true;
                    while (moveIterator)
                    {
                        moveIterator = edgeFIter.MoveNext();

                        if (!moveIterator) break;

                        EdgeArray edges = edgeFIter.Current as EdgeArray;
                        foreach (Edge edge in edges)
                        {
                            Curve curve = edge.AsCurve();
                            if (curve is Arc)
                            {
                                ohe_Shape = OpenigHoleShape.Round;
                                moveIterator = false;
                                break;
                            }
                        }
                    }

                    // Получаю ширину и высоту
                    double[] widthAndHeight = GeometryWorker.GetSolidWidhtAndHeight_ByDirection(iosElemEnt.ARKRIOS_IntesectionSolid, hostDir);
                    double resultWidth = widthAndHeight[0];
                    double resultHeight = widthAndHeight[1];
                    double resultRadius = widthAndHeight[1];


                    // Получаю координаты центра пересечения
                    XYZ docCoord = iosElemEnt.ARKRIOS_IntesectionSolid.ComputeCentroid();


                    // Создаю сущность AROpeningHoleEntity
                    AROpeningHoleEntity arEntity = new AROpeningHoleEntity(ohe_Shape, MainDBService.Get_DBDocumentSubDepartment(iosElemEnt.IOS_LinkDocument).Code, arkrElemEnt.ARKRHost_Element, docCoord)
                        .SetFamilyPathAndName(doc)
                        as AROpeningHoleEntity;

                    arEntity.SetGeomParams();
                    arEntity.SetGeomParamsRoundData(resultHeight, resultWidth, resultRadius, OpenHoleExpandedValue);
                    arEntity.UpdatePointData_ByShape();

                    result.Add(arEntity);

                    ++progressInfoViewModel.CurrentProgress;
                    progressInfoViewModel.DoEvents();
                }
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
                if (arOHEColl == null || CheckWSAvailableError(doc, arOHEColl.Select(arOHE => arOHE.OHE_Element.Id)))
                    return;

                ProgressInfoViewModel progressInfoViewModel = new ProgressInfoViewModel();
                ProgressWindow window = new ProgressWindow(progressInfoViewModel);
                window.Show();

                AROpeningHoleEntity[] unionAR_OHE = AROpeningHoleEntity.ClearCollectionByJoinedHosts(doc, AROpeningHoleEntity.CreateUnionOpeningHole(doc, arOHEColl));
                if (unionAR_OHE != null)
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new AR_OHE_MakerWithUnion(unionAR_OHE, "KPLN: Объединить отверстия", this, true, progressInfoViewModel));
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
                    "Выберите отверстия АР/КЖ, которые нужно объеденить в одно");
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

                AROpeningHoleEntity arOHE = new AROpeningHoleEntity(Core.MainEntity.OpenigHoleShape.Rectangular, fi.Symbol.Name, hostElem, locPnt.Point, selectedElem);

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

                ProgressInfoViewModel progressInfoViewModel = new ProgressInfoViewModel();
                ProgressWindow window = new ProgressWindow(progressInfoViewModel);
                window.Show();

                AROpeningHoleEntity[] arEntities = GetAROpeningsFromIOSTask(doc, iosTasks);

                if (arEntities.Any())
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new AR_OHE_Maker(arEntities, "KPLN: Отверстия по заданию", progressInfoViewModel));
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
                    //LocationPoint locPnt = linkedElement.Location as LocationPoint;
                    XYZ locPnt = GeometryWorker.GetRevitElemSolid(linkedElement).ComputeCentroid();
                    IOSOpeningHoleTaskEntity iosTask = new IOSOpeningHoleTaskEntity(linkedDoc, linkedElement, locPnt)
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
                .Where(e => ARKRElemsWorker.ElemCatLogicalOrFilter.PassesFilter(e) && ARKRElemsWorker.ARKRElemExtraFilterFunc(e)));

            result.AddRange(allHostFromDocumentColl
                .Where(e => ARKRElemsWorker.ElemCatLogicalOrFilter.PassesFilter(e) && ARKRElemsWorker.ARKRElemExtraFilterFunc(e)));

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
