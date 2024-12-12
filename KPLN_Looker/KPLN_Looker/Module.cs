﻿using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Library_Bitrix24Worker;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Loader.Common;
using KPLN_Looker.ExecutableCommand;
using KPLN_Looker.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Looker
{
    public class Module : IExternalModule
    {
        /// <summary>
        /// Общий фильтр (для анализа на копирование эл-в)
        /// </summary>
        private static LogicalOrFilter _resultFamInstFilter;
        /// <summary>
        /// Метка пользователя ИОС (для анализа на коллизии)
        /// </summary>
        private static bool _isIOSUser = false;
        /// <summary>
        /// Метка закрытого для пользователя проекта
        /// </summary>
        private bool _isProjectCloseToUser;

        public Module()
        {
            ModuleDBWorkerService = new DBWorkerService();
            IntersectedElemFinderService.IntersecedServiceDBWorkerService = new DBWorkerService();
        }
        public static DBWorkerService ModuleDBWorkerService { get; private set; }

        /// <summary>
        /// Указатель на окно ревит
        /// </summary>
        public static IntPtr MainWindowHandle { get; private set; }

        public Result Close()
        {
            return Result.Succeeded;
        }

        public Result Execute(UIControlledApplication application, string tabName)
        {
            MainWindowHandle = application.MainWindowHandle;

            try
            {
                // Перезапись ini-файла
                INIFileService iNIFileService = new INIFileService(ModuleDBWorkerService.CurrentDBUser, application.ControlledApplication.VersionNumber);
                if (!iNIFileService.OverwriteINIFile())
                {
                    throw new Exception($"Ошибка при перезаписи ini-файла");
                }

                //Подписка на события
                application.ControlledApplication.ApplicationInitialized += OnApplicationInitialized;
                application.ControlledApplication.DocumentOpened += OnDocumentOpened;
                application.ViewActivated += OnViewActivated;
                application.ControlledApplication.DocumentChanged += OnDocumentChanged;
                application.ControlledApplication.DocumentSynchronizedWithCentral += OnDocumentSynchronized;
#if Debug2020 || Debug2023
                application.ControlledApplication.FamilyLoadingIntoDocument += OnFamilyLoadingIntoDocument;
#else
                if (!ModuleDBWorkerService.CurrentDBUserSubDepartment.Code.ToUpper().Contains("BIM"))
                    application.ControlledApplication.FamilyLoadingIntoDocument += OnFamilyLoadingIntoDocument;
#endif
                if (ModuleDBWorkerService.CurrentDBUserSubDepartment.Code.ToUpper().Contains("ОВиК")
                    || ModuleDBWorkerService.CurrentDBUserSubDepartment.Code.ToUpper().Contains("ВК")
                    || ModuleDBWorkerService.CurrentDBUserSubDepartment.Code.ToUpper().Contains("ЭОМ")
                    || ModuleDBWorkerService.CurrentDBUserSubDepartment.Code.ToUpper().Contains("СС"))
                    _isIOSUser = true;

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                Print($"Ошибка: {ex.Message}", MessageType.Error);
                return Result.Failed;
            }
        }

        /// <summary>
        /// Событие, которое будет выполнено при инициализации приложения
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void OnApplicationInitialized(object sender, ApplicationInitializedEventArgs e)
        {
            if (ModuleDBWorkerService.CurrentDBUser.IsUserRestricted)
            {
                MessageBox.Show(
                    $"Ваша работа ограничена работой в тестовых файлах. " +
                    $"Любой факт попытки открытия/синхронизации при работе с файлами с диска " +
                    $"Y:\\ - будет передан в BIM-отдел",
                    "Предупреждение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Asterisk);
            }
        }

        /// <summary>
        /// Событие на активацию вида
        /// </summary>
        private static void OnViewActivated(object sender, ViewActivatedEventArgs args)
        {
            Autodesk.Revit.DB.View activeView = args.CurrentActiveView;
            Document doc = args.Document;
            UIDocument uidoc = new UIDocument(doc);

#if Revit2020 || Revit2023
            // Игнор НЕ мониторинговых моделей
            if (MonitoredDocFilePath(doc) == null)
                return;
#endif

            #region Обновление кэша по сервисам
#if Debug2020 || Debug2023
            IntersectedElemFinderService.UpdateDataByCurrentDocument(doc);
            IntersectedElemFinderService.UpdateIntersectElemLinkEntitiesByCurrentView(doc);
#else
            if (_isIOSUser)
            {
                IntersectedElemFinderService.UpdateDataByCurrentDocument(doc);
                IntersectedElemFinderService.UpdateIntersectElemLinkEntitiesByCurrentView(doc);
            }
#endif
            #endregion

            #region Закрываю вид, если он для бим-отдела
            if (!(activeView is View3D _)
                || (!activeView.Title.ToUpper().Contains("BIM360")
                    && !activeView.Title.ToUpper().Contains("NAVISWORKS")
                    && !activeView.Title.ToUpper().Contains("GSTATION")
                    && !activeView.Title.ToUpper().Contains("NWC")
                    && !activeView.Title.ToUpper().Contains("NWD"))
                || ModuleDBWorkerService.CurrentDBUserSubDepartment.Code.ToUpper().Contains("BIM"))
                return;

            IList<UIView> openViews = uidoc.GetOpenUIViews();

            TaskDialog td;
            if (openViews.Count > 1)
            {
                td = new TaskDialog("ВНИМАНИЕ")
                {
                    MainIcon = TaskDialogIcon.TaskDialogIconError,
                    MainInstruction = "Данный вид предназначен только для bim-отдела. Его запрещено открывать или редактировать, поэтому он зароется",
                    CommonButtons = TaskDialogCommonButtons.Ok,
                };

                KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new ViewCloser(activeView.Id));
            }
            else
            {
                td = new TaskDialog("ВНИМАНИЕ: Закройте вид!")
                {
                    MainIcon = TaskDialogIcon.TaskDialogIconError,
                    MainInstruction = "Данный вид предназначен только для bim-отдела. Его запрещено открывать или редактировать. Вид нужно закрыть",
                    CommonButtons = TaskDialogCommonButtons.Ok,
                };
            }

            td?.Show();
            #endregion
        }

        /// <summary>
        /// Событие на изменение в документе
        /// </summary>
        private static void OnDocumentChanged(object sender, DocumentChangedEventArgs args)
        {
#if Debug2020 || Debug2023
            if (IntersectedElemFinderService.IsDocumentAnalyzing)
            {
                string transName = args.GetTransactionNames().FirstOrDefault();
                if (transName.Equals("Загрузить связь"))
                    IntersectedElemFinderService.UpdateIntersectElemLinkEntitiesByCurrentView(args.GetDocument());

                IntersectedElemFinderService.UpdateIntersectElemDocEntities(args.GetDocument());
                // !!! ТУТ НЕ ХВАТАЕТ РЕАЛИЗАЦИИ УДАЛЕНИЯ УЖЕ РАЗМЕЩЕННЫХ ЭКЗЕМПЛЯРОВ. СУТЬ УДАЛЕНИЯ - КАК У GC !!!

                CheckError_IOSModelIsIntersected(args);
            }

            CheckError_FamilyCopiedFromOtherFile(args);
        }
#else
            // Фильтрация по имени проекта
            string docPath = MonitoredDocFilePath(args.GetDocument());
            if (docPath != null)
            {
                // Анализ загрузки семейств путем копирования
                if (// Отлов проекта ПШМ1.1_РД_ОВ. Делает субчик на нашем компе. Семейства правит сам.
                    !docPath.Contains("Жилые здания\\Пушкино, Маяковского, 1 очередь\\10.Стадия_Р\\7.4.ОВ\\")
                    // Отлов проекта Школа 825. Его дорабатываем за другой организацией
                    && !docPath.ToLower().Contains("sh1-"))
                    CheckError_FamilyCopiedFromOtherFile(args);

                // Анализ моделей ИОС на генерацию коллизий (!!!ЗАГЛУШЕНО ДО ТЕСТИРОВАНИЯ!!!)
                //string transName = args.GetTransactionNames().FirstOrDefault();
                //if (transName.Equals("Загрузить связь"))
                //    IntersectedElemFinderService.UpdateIntersectElemLinkEntitiesByCurrentView(args.GetDocument());

                //IntersectedElemFinderService.UpdateIntersectElemDocEntities(args.GetDocument());

                //CheckError_IOSModelIsIntersected(args);
            }
        }
#endif

    /// <summary>
    /// Контроль процесса загрузки семейств в проекты КПЛН
    /// </summary>
    private static void OnFamilyLoadingIntoDocument(object sender, FamilyLoadingIntoDocumentEventArgs args)
    {
        DocumentSet appDocsSet;
        if (sender is Autodesk.Revit.ApplicationServices.Application app)
            appDocsSet = app.Documents;
        else
            return;

        Document prjDoc = args.Document;
        string familyName = args.FamilyName;
        string familyPath = args.FamilyPath;

        string docPath = MonitoredDocFilePath(prjDoc);
        if (docPath == null
            // Отлов проекта ПШМ1.1_РД_ОВ. Делает субчик на нашем компе. Семейства правит сам.
            || docPath.Contains("Жилые здания\\Пушкино, Маяковского, 1 очередь\\10.Стадия_Р\\7.4.ОВ\\")
            // Отлов проекта Школа 825. Его дорабатываем за другой организацией
            || docPath.ToLower().Contains("sh1-"))
            return;

        foreach (Document doc in appDocsSet)
        {
            if (doc.IsLinked)
                continue;

            if (!doc.Title.Contains($"{familyName}"))
                continue;

            if (doc.IsFamilyDocument)
            {
                Family family = doc.OwnerFamily;
                Category famCat = family.FamilyCategory;
                BuiltInCategory bic = (BuiltInCategory)famCat.Id.IntegerValue;

                if (!IsFamilyMonitoredError(prjDoc, bic, familyName, familyPath))
                    return;
            }
            else
                throw new Exception("Ошибка определения типа файла. Обратись к разработчику!");
        }

        UserVerify userVerify = new UserVerify("[BEP]: Загружать семейства можно только с диска X (из папки проекта, если она есть)");
        userVerify.ShowDialog();

        switch (userVerify.Status)
        {
            case UIStatus.RunStatus.CloseBecauseError:
                TaskDialog.Show("Запрещено", "Не верный пароль, в загрузке семейства отказано!");
                args.Cancel();
                break;
            case UIStatus.RunStatus.Close:
                args.Cancel();
                break;
        }
    }

    /// <summary>
    /// Выдача имени файла с проверкой на необходимость в контроле действий
    /// </summary>
    private static string MonitoredDocFilePath(Document doc)
    {
        string fileName = doc.IsWorkshared
            ? ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath())
            : doc.PathName;

        if (doc.IsWorkshared
            && !doc.IsDetached
            && !doc.IsFamilyDocument
            && (fileName.ToLower().Contains("stinproject.local\\project\\") || fileName.ToLower().Contains("rsn"))
            && !fileName.EndsWith("rte")
            && !fileName.ToLower().Contains("\\lib\\")
            && !fileName.ToLower().Contains("конц")
            && !fileName.ToLower().Contains("kon"))
            return fileName;

        return null;
    }

    /// <summary>
    /// Проверка семейства на наличие ошибки источника
    /// </summary>
    /// <param name="doc">Рeвит документ</param>
    /// <param name="bic">BuiltInCategory семейства</param>
    /// <param name="familyName">Имя семейства</param>
    /// <param name="familyPath">Путь к семейству</param>
    /// <returns></returns>
    private static bool IsFamilyMonitoredError(
        Document doc,
        BuiltInCategory bic,
        string familyName,
        string familyPath = null)
    {
        // Глобальный отлов по пути семейства (если оно задано).
        // Уточнение для ЛОКАЛЬНЫХ проектов
        if (!string.IsNullOrEmpty(familyPath)
            && (familyPath.StartsWith("X:\\BIM\\3_Семейства") || familyPath.Contains("KPLN_Loader"))
            && (doc.Title.Contains("СЕТ_1") && familyPath.StartsWith("X:\\BIM\\3_Семейства\\8_Библиотека семейств Самолета")))
            return false;
        // Игнор локальных проектов. Для СЕТ плохой пример, на будущее - лучше библиотеку под проект выносить в другой корень, иначе это усложняет анализ
        else if (!string.IsNullOrEmpty(familyPath)
            && (familyPath.StartsWith("X:\\BIM\\3_Семейства") || familyPath.Contains("KPLN_Loader"))
            && (!doc.Title.Contains("СЕТ_1") && !familyPath.StartsWith("X:\\BIM\\3_Семейства\\8_Библиотека семейств Самолета")))
            return false;

        #region Игнорирую семейства, которые могут редактировать проектировщики
        // Отлов семейств марок (могут разрабатывать все)
        if (bic.Equals(BuiltInCategory.OST_ProfileFamilies)
            || bic.Equals(BuiltInCategory.OST_DetailComponents)
            || bic.Equals(BuiltInCategory.OST_GenericAnnotation)
            || bic.Equals(BuiltInCategory.OST_DetailComponentsHiddenLines)
            || bic.Equals(BuiltInCategory.OST_DetailComponentTags))
            return false;

        Category famCat = Category.GetCategory(doc, bic);
        // Отлов семейств марок (могут разрабатывать все), за исключением штампов, подписей и жуков
        if (famCat.CategoryType.Equals(CategoryType.Annotation)
            && !familyName.StartsWith("020_")
            && !familyName.StartsWith("022_")
            && !familyName.StartsWith("023_")
            && !familyName.ToLower().Contains("жук"))
            return false;

        // Отлов семейств по категории
        if (
            // Лестничные марши и площадки, которые по форме зависят от проектов (могут разрабатывать все)
            (bic.Equals(BuiltInCategory.OST_GenericModel) && (familyName.StartsWith("208_") || familyName.StartsWith("209_")))
            // Семейства ограждений (со всеми вложенными эл-тами), которые по форме зависят от проектов (могут разрабатывать все)
            || (bic.Equals(BuiltInCategory.OST_StairsRailing) || bic.Equals(BuiltInCategory.OST_StairsRailingBaluster))
            // Семейства соед. Деталей каб. Лотков производителей: Ostec, Dkc
            || (bic.Equals(BuiltInCategory.OST_CableTrayFitting) && (familyName.ToLower().Contains("ostec") || familyName.ToLower().Contains("dkc"))))
            return false;

        // Отлов семейств ферм, которые по форме зависят от проектов (могут разрабатывать КР)
        if ((ModuleDBWorkerService.CurrentDBUserSubDepartment.Code.ToUpper().Contains("BIM")
             || ModuleDBWorkerService.CurrentDBUserSubDepartment.Code.ToUpper().Contains("КР"))
            && bic.Equals(BuiltInCategory.OST_Truss))
            return false;
        #endregion

        return true;
    }

    /// <summary>
    /// Поиск коллизий ИОС на ИОС при попытке добавлять линейные элементы в моделях ВИС
    /// </summary>
    private static void CheckError_IOSModelIsIntersected(DocumentChangedEventArgs args)
    {
        Document doc = args.GetDocument();

        List<Element> addedLinearElems = args
            .GetAddedElementIds()
            .Select(id => doc.GetElement(id))
            .Where(el => el.Category != null && IntersectedElemFinderService.BuiltInCatIDs.Any(bicId => el.Category.Id.IntegerValue == bicId))
            .ToList();

        List<Element> modifyedLinearElems = args
            .GetModifiedElementIds()
            .Select(id => doc.GetElement(id))
            .Where(el => el.Category != null && IntersectedElemFinderService.BuiltInCatIDs.Any(bicId => el.Category.Id.IntegerValue == bicId))
            .ToList();

        if (!modifyedLinearElems.Any() && !addedLinearElems.Any())
            return;

        addedLinearElems.AddRange(modifyedLinearElems);

        List<XYZ> intersectedPoints = new List<XYZ>();
        // Тут и далее хардкод по разделам КПЛН для экономии ресурсов
        switch (IntersectedElemFinderService.CurrentDocDBSubDepartmentId)
        {
            case -1:
                goto case 8;
            case 1:
                goto case 8;
            case 2:
                goto case 8;
            case 3:
                goto case 8;
            case 4:
                IntersectedElemFinderService.CheckIf_OVLoad = true;
                goto case 7;
            case 5:
                IntersectedElemFinderService.CheckIf_VKLoad = true;
                goto case 7;
            case 6:
                IntersectedElemFinderService.CheckIf_EOMLoad = true;
                goto case 7;
            case 7:
                IntersectedElemFinderService.CheckIf_SSLoad = true;
                intersectedPoints = IntersectedElemFinderService.GetIntersectPoints(addedLinearElems);
                break;
            case 8:
                return;
        }

        bool isAllLinkLoad = IntersectedElemFinderService.CheckIf_OVLoad
            && IntersectedElemFinderService.CheckIf_VKLoad
            && IntersectedElemFinderService.CheckIf_EOMLoad &&
            IntersectedElemFinderService.CheckIf_SSLoad;

        // ВСЕ ИОС связи должны быть загружены
        if (!(isAllLinkLoad))
        {
            TaskDialog td = new TaskDialog("ВНИМАНИЕ: Вы не открыли связи!")
            {
                MainIcon = TaskDialogIcon.TaskDialogIconError,
                MainInstruction = "При моделировании линейных эл-в ВИС - необходимо подгружать ВСЕ связи ВИС, и учитывать их. Анализ на коллизии проведен НЕ в полном объеме",
                CommonButtons = TaskDialogCommonButtons.Ok,
            };

            td?.Show();
        }

        if (intersectedPoints.Any())
        {
            TaskDialog td = new TaskDialog("ВНИМАНИЕ: Вы сгенерировали коллизии!")
            {
                MainIcon = TaskDialogIcon.TaskDialogIconError,
                MainInstruction = "При моделировании линейной сети вы создали коллизии с другими сетями. ЗАВЕРШИТЕ создание, чтобы получить точки пересечений",
                CommonButtons = TaskDialogCommonButtons.Ok,
            };

            td?.Show();

            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new IntersectPointMaker(intersectedPoints));
        }
    }

    /// <summary>
    /// Поиск ошибки при попытке скопировать семейство (FamilyInstance) из другого проекта через буфер обмена
    /// </summary>
    private static void CheckError_FamilyCopiedFromOtherFile(DocumentChangedEventArgs args)
    {
        string transName = args.GetTransactionNames().FirstOrDefault();
        if (transName == null)
            return;

        string lowerTransName = transName.ToLower();
        if (!lowerTransName.Contains("встав") && !lowerTransName.Contains("paste"))
            return;

        Document doc = args.GetDocument();

        // Коллекция добавленых анализируемых FamilyInstance
        SetResultFamInstFilterForAddedElems();
        FamilyInstance[] monitoredFamInsts = args
            .GetAddedElementIds(_resultFamInstFilter)
            .Select(id => doc.GetElement(id))
            .Cast<FamilyInstance>()
            .Where(fi => IsFamilyMonitoredError(doc, (BuiltInCategory)fi.Category.Id.IntegerValue, fi.Symbol.FamilyName))
            .ToArray();

        if (!monitoredFamInsts.Any())
            return;

        // Есть возможность копировать листы через буфер обмена. Ревит автоматом добавляет инкременту номеру, имя остаётся тем же
        if (monitoredFamInsts.All(fi => fi.Category.Id.IntegerValue == (int)BuiltInCategory.OST_TitleBlocks))
            return;

        string docCentralPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath());
        DBProject docDBProject = ModuleDBWorkerService.Get_DBProjectByRevitDocFile(docCentralPath);
        if (docDBProject == null)
            return;

        string docTitle = doc.Title;
        int countDifferenceProjects = 0;
        DocumentSet docSet = doc.Application.Documents;
        foreach (Document openDoc in docSet)
        {
            // Если это линк - ок, пусть копируют
            if (openDoc.IsLinked || openDoc.Title.Equals(docTitle))
                continue;

            if (openDoc.IsWorkshared)
            {
                string openDocCentralPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(openDoc.GetWorksharingCentralModelPath());
                DBProject openDocDBProject = ModuleDBWorkerService.Get_DBProjectByRevitDocFile(openDocCentralPath);
                // Если проекты из БД отличаются - блокирую
                if (openDocDBProject != null && docDBProject.Id == openDocDBProject.Id)
                    continue;
            }

            countDifferenceProjects++;
        }

        if (countDifferenceProjects > 0)
        {
            UserVerify userVerify = new UserVerify(
                "[ОШИБКА]: Выявлена попытка копирования эл-в через буфер обмена. При несколько открытых моделях РАЗНЫХ, или НЕОПРЕДЕЛЕННЫХ проектов в одном Revit - данный функционал запрещен.\n" +
                "[ЧТО ДЕЛАТЬ]: Если нужно скопировать фрагмент из другой модели - обратитесь в BIM-отдел. Иначе - оставьте открытым только ОДИН проект.\n" +
                "[ИНФО]: Вместо буфера обмена - пользуйтесь командами со вкладки \"Изменить\"->\"Изменить\"->\"Копировать\").");
            userVerify.ShowDialog();

            switch (userVerify.Status)
            {
                case UIStatus.RunStatus.CloseBecauseError:
                    TaskDialog.Show("Запрещено", "Не верный пароль, отказано!");
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new UndoEvantHandler());
                    break;
                case UIStatus.RunStatus.Close:
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new UndoEvantHandler());
                    break;
            }
        }
    }

    /// <summary>
    /// Метод кэширования и создания фильтра для фильтрации эл-в при вставке в ревит
    /// </summary>
    private static void SetResultFamInstFilterForAddedElems()
    {
        if (_resultFamInstFilter != null)
            return;

        ElementClassFilter classFilter = new ElementClassFilter(typeof(FamilyInstance));

        _resultFamInstFilter = new LogicalOrFilter(new List<ElementFilter>()
            {
                classFilter,
            });
    }

    /// <summary>
    /// У РС нет возможности откатиться. Делаю свой бэкап на наш сервак
    /// </summary>
    /// <param name="doc">Ревит док</param>
    /// <param name="pathTo">Путь, куда делать бэкап</param>
    private static void RSBackupFile(Document doc, string pathTo)
    {
        string copyTaskMsg = string.Empty;
        Task copyLocalFileTask = Task.Run(() =>
        {
            FileInfo localFI = new FileInfo(doc.PathName);
            if (localFI.Exists)
            {
                if (Directory.Exists(pathTo))
                {
                    string archCopyPath = $"{pathTo}\\{doc.Title}_{DateTime.Now:MM_d_H_m}.rvt";
                    FileInfo archCopyFI = localFI.CopyTo(archCopyPath);
                    if (!archCopyFI.Exists)
                        copyTaskMsg = "Не удалось сделать архивную копию. Обратись в BIM-отдел!";
                }
            }
            else
                copyTaskMsg = "Не удалось найти локальную копию. Обратись в BIM-отдел!";
        });

        string clearTaskMsg = string.Empty;
        Task clearArchCopyFilesTask = Task.Run(() =>
        {
            // Лимит на архивные копии для конкретного файла
            const int archFilesLimit = 10;
            string[] namePrepareArr = doc
                .Title
                .Split(new[] { $"_{ModuleDBWorkerService.CurrentDBUser.SystemName}" }, StringSplitOptions.None);

            if (namePrepareArr.Length == 0)
                clearTaskMsg = $"Не удалось определить имя модели из хранилища " +
                               $"для пользователя {ModuleDBWorkerService.CurrentDBUser.SystemName}. Обратись в BIM-отдел!";
            else
            {
                string centralFileName = namePrepareArr[0];
                if (!Directory.Exists(pathTo))
                    return;

                string[] archFiles = Directory.GetFiles(pathTo);
                FileInfo[] currentCentralArchCopies = archFiles
                    .Where(a => a.Contains(centralFileName))
                    .Select(a => new FileInfo(a))
                    .OrderBy(fi => fi.CreationTime)
                    .ToArray();

                if (currentCentralArchCopies.Count() <= archFilesLimit)
                    return;

                int startCount = currentCentralArchCopies.Count() - archFilesLimit;
                while (startCount > 0)
                {
                    startCount--;
                    FileInfo archCopyToDel = currentCentralArchCopies[startCount];
                    archCopyToDel.Delete();
                }
            }
        });
        Task.WaitAll(new Task[2] { copyLocalFileTask, clearArchCopyFilesTask });

        if (copyTaskMsg != string.Empty)
            Print($"Ошибка при копировании резервного файла: {copyTaskMsg}", MessageType.Error);

        if (clearTaskMsg != string.Empty)
            Print($"Ошибка при очистке старых резервных копий: {clearTaskMsg}", MessageType.Error);
    }

    /// <summary>
    /// Событие на открытый документ
    /// </summary>
    private void OnDocumentOpened(object sender, DocumentOpenedEventArgs args)
    {
        Document doc = args.Document;
        if (MonitoredDocFilePath(doc) == null)
            return;

        string centralPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath());
        #region Отлов пользователей с ограничением допуска к работе ВО ВСЕХ ПРОЕКТАХ
        if (ModuleDBWorkerService.CurrentDBUser.IsUserRestricted)
        {
            BitrixMessageSender.SendMsg_ToBIMChat(
                $"Сотрудник: {ModuleDBWorkerService.CurrentDBUser.Surname} {ModuleDBWorkerService.CurrentDBUser.Name} " +
                $"из отдела {ModuleDBWorkerService.CurrentDBUserSubDepartment.Code}\n" +
                $"Статус допуска: Ограничен в работе с реальными проектами " +
                $"(IsUserRestricted={ModuleDBWorkerService.CurrentDBUser.IsUserRestricted})\n" +
                $"Действие: Открыл файл {doc.Title}.");

            MessageBox.Show(
                $"Вы открыли проект с диска Y:\\. Напомню - Ваша работа ограничена тестовыми файлами! " +
                $"Данные переданы в BIM-отдел",
                "KPLN: Ошибка",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
        #endregion

        #region Обработка проектов КПЛН
        DBProject dBProject = ModuleDBWorkerService.Get_DBProjectByRevitDocFile(centralPath);
        if (dBProject != null)
        {
            // Ищу документ
            int prjDBSubDepartmentId = ModuleDBWorkerService.Get_DBDocumentSubDepartmentId(doc);
            DBDocument dBDocument = ModuleDBWorkerService.Get_DBDocumentByRevitDocPathAndDBProject(centralPath, dBProject, prjDBSubDepartmentId);
            if (dBDocument == null)
            {
                // Создаю, если не нашел
                DBSubDepartment dBSubDepartment = ModuleDBWorkerService.CurrentDBUserSubDepartment;
                DBUser dbUser = ModuleDBWorkerService.CurrentDBUser;
                dBDocument = ModuleDBWorkerService.Create_DBDocument(
                    centralPath,
                    dBProject.Id,
                    prjDBSubDepartmentId,
                    dbUser.Id,
                    ModuleDBWorkerService.CurrentTimeForDB(),
                    false);
            }

            //Обрабатываю документ
            if (dBDocument == null)
            {
                // Вывожу окно, если документ не связан с проектом из БД
                TaskDialog td = new TaskDialog("ОШИБКА")
                {
                    MainIcon = TaskDialogIcon.TaskDialogIconError,
                    MainInstruction = "Не удалось определить документ в БД. Скинь скрин в BIM-отдел",
                    FooterText = $"Специалисту BIM-отдела: файл - {centralPath}",
                    CommonButtons = TaskDialogCommonButtons.Ok,
                };
                td.Show();
            }
            else
            {
                ModuleDBWorkerService.Update_DBDocumentIsClosedStatus(dBProject);
                DBProjectMatrix[] currentPrjMatrixColl = ModuleDBWorkerService.CurrentDBProjectMatrixColl.Where(prj => dBProject.Id == prj.ProjectId).ToArray();

                // Вывожу окно, если документ ЗАКРЫТ к редактированию
                if (dBProject.IsClosed)
                {
                    MessageBox.Show(
                        "Вы открыли ЗАКРЫТЫЙ проект. Работа в нём запрещена!\nЧтобы получить доступ на " +
                        "внесение изменений в этот проект - обратитесь в BIM-отдел.\n" +
                        "Чтобы открыть проект для ознакомления - откройте его с ОТСОЕДИНЕНИЕМ" +
                        "\nИНФО: Сейчас файл закроется",
                        "KPLN: Закрытый проект",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);

                    #region Извещение в чат bim-отдела
                    BitrixMessageSender.SendMsg_ToBIMChat(
                        $"Сотрудник: {ModuleDBWorkerService.CurrentDBUser.Surname} {ModuleDBWorkerService.CurrentDBUser.Name} из отдела {ModuleDBWorkerService.CurrentDBUserSubDepartment.Code}\n" +
                        $"Статус допуска: Сотрудник открыл ЗАКРЫТЫЙ проект\n" +
                        $"Действие: Открыл файл {doc.Title}.");
                    #endregion


                    #region Извещение пользователю как делать правильно
                    int currentUserBitrixId = ModuleDBWorkerService.CurrentDBUser.BitrixUserID;
                    if (currentUserBitrixId != -1)
                    {
                        string jsonRequestToUser = $@"{{
                                    ""DIALOG_ID"": ""{currentUserBitrixId}"",
                                    ""MESSAGE"": ""Стадия проекта {doc.Title} закрыта. Вы попытались открыть [b]закрытый проект[/b]. Если нужно открыть проект с целью просмотра (обучение, анализ и т.п.), то нужно это делать с [b]отсоединением[/b]"",
                                    ""ATTACH"": [
                                        {{
                                            ""IMAGE"": {{
                                                ""NAME"": ""KPLN_Looker_CentralFile_Open.jpg"",
                                                ""LINK"": ""https://kpln.bitrix24.ru/disk/showFile/1783104/?&ncc=1&ts=1729584883&filename=KPLN_Looker_CentralFile_Open.jpg"",
                                                ""PREVIEW"": ""https://kpln.bitrix24.ru/disk/showFile/1783104/?&ncc=1&ts=1729584883&filename=KPLN_Looker_CentralFile_Open.jpg"",
                                                ""WIDTH"": ""1000"",
                                                ""HEIGHT"": ""1000""
                                            }}
                                        }}
                                    ]
                                }}";

                        BitrixMessageSender.SendMsg_ToUser_ByWebhookKeyANDJSONRequest(
                            "https://kpln.bitrix24.ru/rest/1310/pzyudfrm0pp3gq19/im.message.add.json",
                            jsonRequestToUser);
                    }
                    #endregion

                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new DocCloser(ModuleDBWorkerService.CurrentDBUser, doc));
                }
                // Отлов пользователей с ограничением допуска к работе в текущем проекте
                else if (currentPrjMatrixColl.Length > 0
                         && currentPrjMatrixColl.All(prj => prj.UserId != ModuleDBWorkerService.CurrentDBUser.Id))
                {
                    _isProjectCloseToUser = true;
                    MessageBox.Show(
                        $"Вы открыли файл проекта {dBProject.Name}. Данный проект идёт с требованиями от заказчика," +
                        $" и с ними необходимо предварительно ознакомиться. Для этого - обратись в BIM-отдел." +
                        $"\nИНФО: Если файл засинхронизировать - он закроется",
                        "KPLN: Ограниченный проект",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);

                    BitrixMessageSender.SendMsg_ToBIMChat(
                        $"Сотрудник: {ModuleDBWorkerService.CurrentDBUser.Surname} {ModuleDBWorkerService.CurrentDBUser.Name} " +
                        $"из отдела {ModuleDBWorkerService.CurrentDBUserSubDepartment.Code}\n" +
                        $"Статус допуска: Данный сотрудник не имеет доступа к этому проекту (его нужно внести в список)\n" +
                        $"Действие: Открыл проект {doc.Title}.");
                }
            }
        }
        else
        {
            TaskDialog td = new TaskDialog("ВНИМАНИЕ")
            {
                MainIcon = TaskDialogIcon.TaskDialogIconInformation,
                MainInstruction = "Вы работаете в незарегистрированном проекте. Скинь скрин в BIM-отдел",
                FooterText = $"Специалисту BIM-отдела: файл - {centralPath}",
                CommonButtons = TaskDialogCommonButtons.Ok,
            };
            td.Show();
        }
        #endregion

    }

    /// <summary>
    /// Событие на синхронизацию файла
    /// </summary>
    private void OnDocumentSynchronized(object sender, DocumentSynchronizedWithCentralEventArgs args)
    {
        Document doc = args.Document;

        if (MonitoredDocFilePath(doc) == null)
            return;

        ModuleDBWorkerService.DropMainCash();

        string centralPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath());
        #region Отлов пользователей с ограничением допуска к работе ВО ВСЕХ ПРОЕКТАХ
        if (ModuleDBWorkerService.CurrentDBUser.IsUserRestricted)
        {
            BitrixMessageSender.SendMsg_ToBIMChat(
                $"Сотрудник: {ModuleDBWorkerService.CurrentDBUser.Surname} {ModuleDBWorkerService.CurrentDBUser.Name} из отдела {ModuleDBWorkerService.CurrentDBUserSubDepartment.Code}\n" +
                $"Статус допуска: Ограничен в работе с реальными проектами (IsUserRestricted={ModuleDBWorkerService.CurrentDBUser.IsUserRestricted})\n" +
                $"Действие: Произвел синхронизацию файла {doc.Title}.");

            MessageBox.Show(
                $"Вы произвели синхронизацию проекта с диска Y:\\, хотя у вас нет к этому доступа (вы не сдали КЛ BIM-отделу). " +
                $"Данные переданы в BIM-отдел и ГИ бюро. Файл будет ЗАКРЫТ.",
                "KPLN: Ошибка",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new DocCloser(ModuleDBWorkerService.CurrentDBUser, doc));
        }
        #endregion

        #region Работа с проектами КПЛН
        DBProject dBProject = ModuleDBWorkerService.Get_DBProjectByRevitDocFile(centralPath);
        if (dBProject == null)
            return;

        int prjDBSubDepartmentId = ModuleDBWorkerService.Get_DBDocumentSubDepartmentId(doc);
        DBDocument dBDocument = ModuleDBWorkerService.Get_DBDocumentByRevitDocPathAndDBProject(centralPath, dBProject, prjDBSubDepartmentId);
        if (dBDocument == null)
            return;

        ModuleDBWorkerService.Update_DBDocumentLastChangedData(dBDocument);

        // Защита закрытого проекта от изменений (файл вообще не должен открываться, но ЕСЛИ это произошло - будет уведомление)
        if (dBDocument.IsClosed)
        {
            BitrixMessageSender.SendMsg_ToBIMChat(
                $"Сотрудник: {ModuleDBWorkerService.CurrentDBUser.Surname} {ModuleDBWorkerService.CurrentDBUser.Name} из отдела {ModuleDBWorkerService.CurrentDBUserSubDepartment.Code}\n" +
                $"Статус допуска: Сотрудник засинхронизировал ЗАКРЫТЫЙ проект (проект все же НЕ удалось закрыть)\n" +
                $"Действие: Произвел синхронизацию в {doc.Title}.");

            MessageBox.Show(
                $"Вы произвели синхронизацию ЗАКРЫТОГО проекта с диска Y:\\. Данные переданы в BIM-отдел. Файл будет ЗАКРЫТ.",
                "KPLN: Ошибка",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new DocCloser(ModuleDBWorkerService.CurrentDBUser, doc));
        }
        // Отлов пользователей с ограничением допуска к работе в текущем проекте
        else if (_isProjectCloseToUser)
        {
            BitrixMessageSender.SendMsg_ToBIMChat(
                $"Сотрудник: {ModuleDBWorkerService.CurrentDBUser.Surname} {ModuleDBWorkerService.CurrentDBUser.Name} из отдела {ModuleDBWorkerService.CurrentDBUserSubDepartment.Code}\n" +
                $"Статус допуска: Данный сотрудник не имеет доступа к этому проекту (его нужно внести в список)\n" +
                $"Действие: Произвел синхронизацию в {doc.Title}.");

            MessageBox.Show(
                $"Вы открыли файл проекта {dBProject.Name}. " +
                $"Данный проект идёт с требованиями от заказчика, и с ними необходимо предварительно ознакомиться. " +
                $"Для этого - обратись в BIM-отдел. Файл будет ЗАКРЫТ.",
                "KPLN: Ошибка",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new DocCloser(ModuleDBWorkerService.CurrentDBUser, doc));
        }
        #endregion

        #region Бэкап версий с RS на наш сервак по проектам
        if (args.Status == RevitAPIEventStatus.Succeeded && dBProject.RevitServerPath != null)
        {


            bool isSET = doc.PathName.Contains("СЕТ_1");
            // Проект Сетунь
            if (isSET)
            {
                if (doc.PathName.Contains("_АР_"))
                    RSBackupFile(doc, "Y:\\Жилые здания\\Самолет Сетунь\\10.Стадия_Р\\5.АР\\1.RVT\\00_Автоархив с Revit-Server");
                else if (doc.PathName.Contains("_КР_"))
                    RSBackupFile(doc, "Y:\\Жилые здания\\Самолет Сетунь\\10.Стадия_Р\\6.КР\\1.RVT\\00_Автоархив с Revit-Server");
                else if (doc.PathName.Contains("_ЭОМ"))
                    RSBackupFile(doc, "Y:\\Жилые здания\\Самолет Сетунь\\10.Стадия_Р\\7.1.ЭОМ\\1.RVT\\00_Автоархив с Revit-Server");
                else if (doc.PathName.Contains("_ВК"))
                    RSBackupFile(doc, "Y:\\Жилые здания\\Самолет Сетунь\\10.Стадия_Р\\7.2.ВК\\1.RVT\\00_Автоархив с Revit-Server");
                else if (doc.PathName.Contains("_ПТ"))
                    RSBackupFile(doc, "Y:\\Жилые здания\\Самолет Сетунь\\10.Стадия_Р\\7.3.АУПТ\\1.RVT\\00_Автоархив с Revit-Server");
                else if (doc.PathName.Contains("_ОВ"))
                    RSBackupFile(doc, "Y:\\Жилые здания\\Самолет Сетунь\\10.Стадия_Р\\7.4.ОВ\\1.RVT\\00_Автоархив с Revit-Server");
                else if (doc.PathName.Contains("_ПБ_") || doc.PathName.Contains("_АК_") || doc.PathName.Contains("_СС_"))
                    RSBackupFile(doc, "Y:\\Жилые здания\\Самолет Сетунь\\10.Стадия_Р\\7.5.СС\\1.RVT\\00_Автоархив с Revit-Server");
            }

            // Проект Матросская тишина
            bool isMTRS = doc.PathName.Contains("МТРС_");
            if (isMTRS)
            {
                if (doc.PathName.Contains("_АР_"))
                    RSBackupFile(doc, "Y:\\Жилые здания\\Матросская Тишина\\10.Стадия_Р\\5.АР\\1.RVT\\00_Автоархив с Revit-Server");
                else if (doc.PathName.Contains("_КР_"))
                    RSBackupFile(doc, "Y:\\Жилые здания\\Матросская Тишина\\10.Стадия_Р\\6.КР\\1.RVT\\00_Автоархив с Revit-Server");
                else if (doc.PathName.Contains("_ЭОМ_"))
                    RSBackupFile(doc, "Y:\\Жилые здания\\Матросская Тишина\\10.Стадия_Р\\7.1.ЭОМ\\1.RVT\\00_Автоархив с Revit-Server");
                else if (doc.PathName.Contains("_ВК_"))
                    RSBackupFile(doc, "Y:\\Жилые здания\\Матросская Тишина\\10.Стадия_Р\\7.2.ВК\\1.RVT\\00_Автоархив с Revit-Server");
                else if (doc.PathName.Contains("_ПТ_"))
                    RSBackupFile(doc, "Y:\\Жилые здания\\Матросская Тишина\\10.Стадия_Р\\7.3.АУПТ\\1.RVT\\00_Автоархив с Revit-Server");
                else if (doc.PathName.Contains("_ОВ_"))
                    RSBackupFile(doc, "Y:\\Жилые здания\\Матросская Тишина\\10.Стадия_Р\\7.4.ОВ\\1.RVT\\00_Автоархив с Revit-Server");
                else if (doc.PathName.Contains("_СС_"))
                    RSBackupFile(doc, "Y:\\Жилые здания\\Матросская Тишина\\10.Стадия_Р\\7.5.СС\\1.RVT\\00_Автоархив с Revit-Server");
            }
        }
        #endregion
    }
}
}
