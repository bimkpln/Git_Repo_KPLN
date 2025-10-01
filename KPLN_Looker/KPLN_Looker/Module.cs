using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Library_Bitrix24Worker;
using KPLN_Library_Forms.UI;
using KPLN_Library_SQLiteWorker;
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
        private static readonly string _webHookUrl = "https://kpln.bitrix24.ru/rest/1310/uemokhg11u78vdvs";

        /// <summary>
        /// Коллекция пользователей БИМ-отдела, которых подписываю на рассылку уведомлений по файлам АР_АФК
        /// </summary>
        private static DBUser[] _arKonFileSubscribersFromBIM;

        /// <summary>
        /// Общий фильтр (для анализа на копирование эл-в)
        /// </summary>
        private static LogicalOrFilter _resultFamInstFilter;

        /// <summary>
        /// Метка времени последнего оповещения
        /// </summary>
        private static DateTime _lastAlarm = new DateTime(2025, 01, 01, 0, 0, 0);

        /// <summary>
        /// Лимит задержки сообщений
        /// </summary>
        private static readonly TimeSpan _delayAlarm = new TimeSpan(0, 10, 0);

        /// <summary>
        /// Метка закрытого для пользователя проекта
        /// </summary>
        private bool _isProjectCloseToUser;

        public Module()
        {
            List<DBUser> bimManager = DBMainService.UserDbService.GetDBUsers_BIMManager().ToList();
            List<DBUser> arBimCoord = DBMainService.UserDbService.GetDBUsers_BIMARCoord().ToList();
            arBimCoord.AddRange(bimManager);

            _arKonFileSubscribersFromBIM = arBimCoord.ToArray();
        }

        /// <summary>
        /// Указатель на окно ревит
        /// </summary>
        public static IntPtr MainWindowHandle { get; private set; }

        public static int RevitVersion { get; private set; }

        /// <summary>
        /// Флаг запуска автопроверок (выключаем из модулей, которые запускают синхрон моделей)
        /// </summary>
        public static bool RunAutoChecks { get; set; } = true;

        public Result Close()
        {
            return Result.Succeeded;
        }

        public Result Execute(UIControlledApplication application, string tabName)
        {
            MainWindowHandle = application.MainWindowHandle;
            RevitVersion = int.Parse(application.ControlledApplication.VersionNumber);

            try
            {
                // Перезапись ini-файла
                INIFileService iNIFileService = new INIFileService(DBMainService.CurrentDBUser, RevitVersion);
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
                application.ControlledApplication.DocumentSaved += OnDocumentSaved;
#else
                if (DBMainService.CurrentUserDBSubDepartment.Code.ToUpper().Contains("АР"))
                    application.ControlledApplication.DocumentSaved += OnDocumentSaved;

                if (!DBMainService.CurrentUserDBSubDepartment.Code.ToUpper().Contains("BIM"))
                    application.ControlledApplication.FamilyLoadingIntoDocument += OnFamilyLoadingIntoDocument;
#endif

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                Print($"Ошибка: {ex.Message}", MessageType.Error);
                return Result.Failed;
            }
        }

        /// <summary>
        /// Получить полное имя открытого файла
        /// </summary>
        /// <param name="doc">Документ Ревит для анализа</param>
        /// <returns></returns>
        public static string GetFileFullName(Document doc) =>
            doc.IsWorkshared && !doc.IsDetached
                ? ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath())
                : doc.PathName;

        /// <summary>
        /// Выдача имени файла с проверкой на необходимость в контроле действий. 
        /// Исключения: модели НЕ с основных серверов КПЛН (диск Y:\\ и любые RS-ы); НЕ шаблоны проектов; не офис КПЛН; НЕ концепции АР
        /// </summary>
        public static string MonitoredDocFilePath_ExceptARKon(Document doc)
        {
            // Такое возможно при работе плагинов с открываением/сохранением моделей (модель не открылась)
            if (doc == null)
                return null;

            string docPathWithARKon = MonitoredDocFilePath(doc);

            if (doc.IsWorkshared
                && !doc.IsDetached
                && docPathWithARKon != null
                && !docPathWithARKon.ToLower().Contains("концепция")
                && !docPathWithARKon.ToLower().Contains("kon_")
                && !docPathWithARKon.Contains(".АГО")
                && !docPathWithARKon.Contains(".АГР"))
                return docPathWithARKon;

            return null;
        }

        /// <summary>
        /// Верхнеуровневая выдача имени файла с проверкой на необходимость в контроле действий
        /// Исключения: модели НЕ с основных серверов КПЛН (диск Y:\\ и любые RS-ы); НЕ шаблоны проектов; не офис КПЛН
        /// </summary>
        public static string MonitoredDocFilePath(Document doc)
        {
            // Такое возможно при работе плагинов с открываением/сохранением моделей (модель не открылась)
            if (doc == null)
                return null;

            string fileFullName = GetFileFullName(doc);

            if (!doc.IsFamilyDocument
#if Revit2020 || Revit2023
                && (fileFullName.ToLower().Contains("stinproject.local\\project\\") || fileFullName.ToLower().Contains("rsn"))
#elif Debug2020 || Debug2023
                && (fileFullName.ToLower().Contains("stinproject.local\\project\\") || fileFullName.ToLower().Contains("fs01\\lib\\отдел bim\\") || fileFullName.ToLower().Contains("rsn"))
#endif
                && !fileFullName.EndsWith("rte")
                // Офис КПЛН
                && !fileFullName.ToLower().Contains("16с13"))
                return fileFullName;

            return null;
        }

        /// <summary>
        /// Событие, которое будет выполнено при инициализации приложения
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void OnApplicationInitialized(object sender, ApplicationInitializedEventArgs e)
        {
            if (DBMainService.CurrentDBUser.IsUserRestricted)
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
            if (MonitoredDocFilePath_ExceptARKon(doc) == null)
                return;
#endif
            #region Закрываю вид, если он для бим-отдела
            if (!(activeView is View3D _)
                || (!activeView.Title.ToUpper().Contains("BIM360")
                    && !activeView.Title.ToUpper().Contains("NAVISWORKS")
                    && !activeView.Title.ToUpper().Contains("GSTATION")
                    && !activeView.Title.ToUpper().Contains("KPLN_NW_")
                    && !activeView.Title.ToUpper().Contains("NWC")
                    && !activeView.Title.ToUpper().Contains("NWD"))
                || DBMainService.CurrentUserDBSubDepartment.Code.ToUpper().Contains("BIM"))
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
            Document document = args.GetDocument();

#if Debug2020 || Debug2023
            // Фильтрация по имени проекта
            string docPath = MonitoredDocFilePath_ExceptARKon(document);
            if (docPath != null)
            {
                CheckAndSendError_FamilyInstanceUserHided(args);

                // Анализ загрузки семейств путем копирования
                if (// Отлов проекта ПШМ1.1_РД_ОВ. Делает субчик на нашем компе. Семейства правит сам.
                    !docPath.Contains("Жилые здания\\Пушкино, Маяковского, 1 очередь\\10.Стадия_Р\\7.4.ОВ\\")
                    // Отлов проекта Школа 825. Его дорабатываем за другой организацией
                    && !docPath.ToLower().Contains("sh1-"))
                    CheckError_FamilyCopiedFromOtherFile(args);
            }
#else
            // Фильтрация по имени проекта
            string docPath = MonitoredDocFilePath_ExceptARKon(document);
            if (docPath != null)
            {
                CheckAndSendError_FamilyInstanceUserHided(args);

                // Анализ загрузки семейств путем копирования
                if (// Отлов проекта ПШМ1.1_РД_ОВ. Делает субчик на нашем компе. Семейства правит сам.
                    !docPath.Contains("Жилые здания\\Пушкино, Маяковского, 1 очередь\\10.Стадия_Р\\7.4.ОВ\\")
                    // Отлов проекта Школа 825. Его дорабатываем за другой организацией
                    && !docPath.ToLower().Contains("sh1-"))
                    CheckError_FamilyCopiedFromOtherFile(args);
            }
#endif
        }

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

            string docPath = MonitoredDocFilePath_ExceptARKon(prjDoc);
            if (docPath == null
                // Отлов проекта ПШМ1.1_РД_ОВ. Делает субчик на нашем компе. Семейства правит сам.
                || docPath.Contains("Жилые здания\\Пушкино, Маяковского, 1 очередь\\10.Стадия_Р\\7.4.ОВ\\")
                // Отлов проекта Школа 825. Его дорабатываем за другой организацией
                || docPath.ToLower().Contains("sh1-"))
                return;

            // Игнор семейств от BimStep
            if (familyName.StartsWith("BS_"))
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
            if (!(bool)userVerify.ShowDialog())
            {
                TaskDialog.Show("Запрещено", "Не верный пароль, в загрузке семейства отказано!");
                args.Cancel();
            }
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
            #region Локальный отлов по пути семейства для проектов
            if (!string.IsNullOrEmpty(familyPath))
            {
                // Уточнение для ЛОКАЛЬНЫХ ПРОЕКТОВ
                bool isSMLT = doc.Title.Contains("СЕТ_1");
                if (isSMLT)
                {
                    if (familyPath.StartsWith("X:\\BIM\\3_Семейства\\8_Библиотека семейств Самолета")
                        || familyPath.Contains("X:\\BIM\\3_Семейства\\0_Общие семейства\\0_Штамп\\022_Подписи в штамп")
                        || familyPath.Contains("X:\\BIM\\3_Семейства\\0_Общие семейства\\0_Штамп\\023_Подписи на титул")
                        || familyPath.Contains("X:\\BIM\\3_Семейства\\0_Общие семейства\\1_Марки\\01_Общие марки для ИОС")
                        || familyPath.Contains("X:\\BIM\\3_Семейства\\0_Общие семейства\\1_Марки\\011_Разрывы скобки")
                        || familyPath.Contains("X:\\BIM\\3_Семейства\\1_АР")
                        || familyPath.Contains("X:\\BIM\\3_Семейства\\2_КР")
                        || familyPath.Contains("Y:\\Жилые здания\\Самолет Сетунь\\4.Оформление\\4.Стадия_Р")
                        || familyPath.Contains("KPLN_Loader"))
                        return false;
                    else
                        return true;
                }

                // Уточнение для ЛОКАЛЬНЫХ СЕМЕЙСТВ
                if (familyPath.StartsWith("X:\\BIM\\3_Семейства\\8_Библиотека семейств Самолета"))
                {
                    if (isSMLT)
                        return false;
                    else
                        return true;
                }
            }
            #endregion


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
            if ((DBMainService.CurrentUserDBSubDepartment.Code.ToUpper().Contains("BIM")
                 || DBMainService.CurrentUserDBSubDepartment.Code.ToUpper().Contains("КР"))
                && bic.Equals(BuiltInCategory.OST_Truss))
                return false;
            #endregion

            // Остальные проекты/семейства
            if (familyPath.StartsWith("X:\\BIM\\3_Семейства") || familyPath.Contains("KPLN_Loader"))
                return false;

            return true;
        }

        /// <summary>
        /// Поиск ошибки скрытия элементов запрещенными действиями. При наличии - отправка пользователю в битрикс
        /// </summary>
        private static void CheckAndSendError_FamilyInstanceUserHided(DocumentChangedEventArgs args)
        {
            string transName = args.GetTransactionNames().FirstOrDefault();
            if (transName == null)
                return;

            string lowerTransName = transName.ToLower();
            if (!lowerTransName.Equals("графика элементов вида")
                && !lowerTransName.Equals("скрыть/изолировать")
                && !lowerTransName.Equals("view specific element graphics")
                && !lowerTransName.Equals("hide/isolate"))
                return;

            // Игнор хардкодом определенных сотрудников, которые внушают доверие
            if (DBMainService.CurrentDBUser.Surname.Equals("Тамарин")
                && DBMainService.CurrentDBUser.Surname.Equals("Егор"))
                return;

            DateTime temp = DateTime.Now;
            if (_delayAlarm < (temp - _lastAlarm))
            {
                _lastAlarm = temp;

                // Отправка уведомления пользователю в Битрикс
                if (DBMainService.CurrentDBUser != null)
                {
                    BitrixMessageSender.SendMsg_ToUser_ByDBUser(
                        DBMainService.CurrentDBUser,
                        "[b][ЗАМЕЧАНИЕ][/b]: Ручное переопределение на виде приводит к сложностям при внесении изменений в модели (новые элементы - уже не соответсвуют ручным правилам).\n" +
                        "[b][ЧТО ДЕЛАТЬ][/b]: В 99% нужно использовать фильтры. Если у тебя есть с этим сложности - напиши BIM-координатору, он поможет.\n");
                }
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

            string fileFullName = GetFileFullName(doc);
            DBProject docDBProject = DBMainService.ProjectDbService.GetDBProject_ByRevitDocFileNameANDRVersion(fileFullName, RevitVersion);
            if (docDBProject == null)
                return;

            string docTitle = doc.Title;
            int countDifferenceProjects = 0;
            DocumentSet docSet = doc.Application.Documents;
            foreach (Document openDoc in docSet)
            {
                // Семейство в игнор (там нечего копировать)
                if (openDoc.IsFamilyDocument)
                    continue;

                // Если это тот же файл - игнор
                if (openDoc.Title.Equals(docTitle))
                    continue;

                // Если это линк - ок, пусть копируют
                if (openDoc.IsLinked)
                    continue;

                // Если проекты из БД отличаются - блокирую
                string openDocFileFullName = GetFileFullName(openDoc);
                DBProject openDocDBProject = DBMainService.ProjectDbService.GetDBProject_ByRevitDocFileNameANDRVersion(openDocFileFullName, RevitVersion);
                if (openDocDBProject != null && docDBProject.Id == openDocDBProject.Id)
                    continue;

                countDifferenceProjects++;
            }

            if (countDifferenceProjects > 0)
            {
                UserVerify userVerify = new UserVerify(
                    "[ОШИБКА]: Выявлена попытка копирования эл-в через буфер обмена. При несколько открытых моделях РАЗНЫХ, или НЕОПРЕДЕЛЕННЫХ (в том числе открытых с ОТСОЕДИНЕНИЕМ) проектов в одном Revit - данный функционал запрещен.\n" +
                    "[ЧТО ДЕЛАТЬ]: Если нужно скопировать фрагмент из другой модели - обратитесь в BIM-отдел. Иначе - оставьте открытым только ОДИН проект.\n" +
                    "[ИНФО]: Вместо буфера обмена - пользуйтесь командами со вкладки \"Изменить\"->\"Изменить\"->\"Копировать\").");

                if (!(bool)userVerify.ShowDialog())
                {
                    TaskDialog.Show("Запрещено", "Не верный пароль, отказано!");
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new UndoEvantHandler());
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
                    .Split(new[] { $"_{DBMainService.CurrentDBUser.SystemName}" }, StringSplitOptions.None);

                if (namePrepareArr.Length == 0)
                    clearTaskMsg = $"Не удалось определить имя модели из хранилища " +
                                   $"для пользователя {DBMainService.CurrentDBUser.SystemName}. Обратись в BIM-отдел!";
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
        /// Анализ файлов Концепций КПЛН и оповещение о создании новых проектов и моделей внутри проектов
        /// </summary>
        private static void ARKonFileSendMsg(Document doc)
        {
            // Если это не концепция АР - в игнор
            if (MonitoredDocFilePath(doc) == null
                || MonitoredDocFilePath_ExceptARKon(doc) != null)
                return;

            // Получаю проект из БД КПЛН
            string fileFullName = GetFileFullName(doc);

            DBProject dBProject = DBMainService.ProjectDbService.GetDBProject_ByRevitDocFileNameANDRVersion(fileFullName, RevitVersion);
            // Проекта нет, а путь принадлежит к мониторинговым ВКЛЮЧАЯ концепции - то оповещение о новом проекте
            if (dBProject == null)
            {
                string centralPathForUser = fileFullName.Replace("\\\\stinproject.local\\project", "Y:");
                foreach (DBUser dbUser in _arKonFileSubscribersFromBIM)
                {
                    BitrixMessageSender.SendMsg_ToUser_ByDBUser(
                        dbUser,
                        $"Сотрудник: {DBMainService.CurrentDBUser.Surname} {DBMainService.CurrentDBUser.Name} " +
                        $"из отдела {DBMainService.CurrentUserDBSubDepartment.Code}\n" +
                        $"Действие: Произвел сохранение/синхронизацию файла незарегестрированного проекта.\n" +
                        $"Имя файла: [b]{doc.Title}[/b].\n" +
                        $"Путь к модели: [b]{centralPathForUser}[/b].",
                        "Y");
                }

                return;
            }

            // Проект есть, но модель еще не зарегестриована в БД - оповещение о новом файле
            DBSubDepartment prjDBSubDepartment = DBMainService.SubDepartmentDbService.GetDBSubDepartment_ByRevitDocFullPath(doc.PathName);
            DBDocument dBDocument = DBDocumentByRevitDocPathAndDBProject(fileFullName, dBProject, prjDBSubDepartment);
            if (dBDocument == null)
            {
                foreach (DBUser dbUser in _arKonFileSubscribersFromBIM)
                {
                    BitrixMessageSender.SendMsg_ToUser_ByDBUser(
                        dbUser,
                        $"Сотрудник: {DBMainService.CurrentDBUser.Surname} {DBMainService.CurrentDBUser.Name} " +
                        $"из отдела {DBMainService.CurrentUserDBSubDepartment.Code}\n" +
                        $"Действие: Произвел сохранение/синхронизацию нового файла проекта [b]{dBProject.Name}_{dBProject.Stage}[/b] (сообщение возникает только при 1м сохранении).\n" +
                        $"Имя файла: [b]{doc.Title}[/b].\n" +
                        $"Путь к модели: [b]{fileFullName}[/b].",
                        "Y");
                }

                // Создаю, если не нашел
                dBDocument = new DBDocument()
                {
                    CentralPath = fileFullName,
                    ProjectId = dBProject.Id,
                    SubDepartmentId = prjDBSubDepartment.Id,
                    LastChangedUserId = DBMainService.CurrentDBUser.Id,
                    LastChangedData = DBMainService.CurrentTimeForDB(),
                    IsClosed = false,
                };

                DBMainService.DocDbService.CreateDBDocument(dBDocument);
            }
        }

        /// <summary>
        /// Получить DBDocument по пути, проекту и id отдела
        /// </summary>
        private static DBDocument DBDocumentByRevitDocPathAndDBProject(string centralPath, DBProject dBProject, DBSubDepartment dBSubDepartment)
        {
            int dBSubDepartmentId = dBSubDepartment == null ? -1 : dBSubDepartment.Id;
            if (dBProject != null)
            {
                IEnumerable<DBDocument> docColl = DBMainService.DocDbService.GetDBDocuments_ByPrjIdAndSubDepId(dBProject.Id, dBSubDepartmentId);
                return docColl.FirstOrDefault(d => d.CentralPath.Equals(centralPath));
            }

            return null;
        }

        /// <summary>
        /// Событие на открытый документ
        /// </summary>
        private void OnDocumentOpened(object sender, DocumentOpenedEventArgs args)
        {
            Document doc = args.Document;
            if (MonitoredDocFilePath_ExceptARKon(doc) == null)
                return;

            #region Отлов пользователей с ограничением допуска к работе ВО ВСЕХ ПРОЕКТАХ
            if (DBMainService.CurrentDBUser.IsUserRestricted)
            {
                BitrixMessageSender.SendMsg_ToBIMChat(
                    $"Сотрудник: {DBMainService.CurrentDBUser.Surname} {DBMainService.CurrentDBUser.Name} " +
                    $"из отдела {DBMainService.CurrentUserDBSubDepartment.Code}\n" +
                    $"Статус допуска: Ограничен в работе с реальными проектами " +
                    $"(IsUserRestricted={DBMainService.CurrentDBUser.IsUserRestricted})\n" +
                    $"Действие: Открыл файл {doc.Title}.");

                MessageBox.Show(
                    $"Вы открыли проект с диска Y:\\. Напомню - Ваша работа ограничена тестовыми файлами! " +
                    $"Данные переданы в BIM-отдел",
                    "KPLN: Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            #endregion

            string fileFullName = GetFileFullName(doc);

            #region Обработка работы в архивных копиях
            if (fileFullName.ToLower().Contains("архив"))
            {
                MessageBox.Show(
                    "Вы открыли АРХИВНЫЙ проект. Работа в нём запрещена, только просмотр!\n" +
                    "\nИНФО: Если попытаетесь что-то синхронизировать - проект закроется",
                    "KPLN: Архивный проект",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                #region Извещение в чат bim-отдела
                BitrixMessageSender.SendMsg_ToBIMChat(
                    $"Сотрудник: {DBMainService.CurrentDBUser.Surname} {DBMainService.CurrentDBUser.Name} из отдела {DBMainService.CurrentUserDBSubDepartment.Code}\n" +
                    $"Статус допуска: Сотрудник открыл АРХИВНЫЙ проект\n" +
                    $"Действие: Открыл файл [b]{doc.Title}[/b].\n" +
                    $"Путь к модели: [b]{fileFullName}[/b].");
                #endregion
            }
            #endregion

            #region Обработка проектов КПЛН
            DBProject dBProject = DBMainService.ProjectDbService.GetDBProject_ByRevitDocFileNameANDRVersion(fileFullName, RevitVersion);
            if (dBProject != null)
            {
                // Ищу документ
                DBSubDepartment prjDBSubDepartment = DBMainService.SubDepartmentDbService.GetDBSubDepartment_ByRevitDocFullPath(doc.PathName);
                DBDocument dBDocument = DBDocumentByRevitDocPathAndDBProject(fileFullName, dBProject, prjDBSubDepartment);
                if (dBDocument == null)
                {
                    dBDocument = new DBDocument()
                    {
                        CentralPath = fileFullName,
                        ProjectId = dBProject.Id,
                        SubDepartmentId = prjDBSubDepartment.Id,
                        LastChangedUserId = DBMainService.CurrentDBUser.Id,
                        LastChangedData = DBMainService.CurrentTimeForDB(),
                        IsClosed = false,
                    };

                    DBMainService.DocDbService.CreateDBDocument(dBDocument);
                }

                //Обрабатываю документ
                DBMainService.DocDbService.UpdateDBDocument_IsClosedByProject(dBProject);

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
                        $"Сотрудник: {DBMainService.CurrentDBUser.Surname} {DBMainService.CurrentDBUser.Name} из отдела {DBMainService.CurrentUserDBSubDepartment.Code}\n" +
                        $"Статус допуска: Сотрудник открыл ЗАКРЫТЫЙ проект\n" +
                        $"Действие: Открыл файл [b]{doc.Title}[/b].\n" +
                        $"Путь к модели: [b]{fileFullName}[/b].");
                    #endregion


                    #region Извещение пользователю как делать правильно
                    int currentUserBitrixId = DBMainService.CurrentDBUser.BitrixUserID;
                    if (currentUserBitrixId != -1)
                    {
                        string jsonRequestToUser = $@"{{
                                    ""DIALOG_ID"": ""{currentUserBitrixId}"",
                                    ""MESSAGE"": ""Проект {doc.Title} закрыт. Актуальный путь - уточняйте у своего руководителя. Вы попытались открыть [b]закрытый проект[/b]. Если нужно открыть проект с целью просмотра (обучение, анализ и т.п.), то нужно это делать с [b]отсоединением[/b]"",
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
                            $"{_webHookUrl}/im.message.add.json",
                            jsonRequestToUser);
                    }
                    #endregion

                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new DocCloser(DBMainService.CurrentDBUser, doc));

                    return;
                }


                // Отлов пользователей с ограничением допуска к работе в текущем проекте
                DBProjectsAccessMatrix[] currentPrjMatrixColl = DBMainService
                    .PrjAccessMatrixDbService
                    .GetDBProjectMatrix_ByProject(dBProject)
                    .ToArray();
                if (currentPrjMatrixColl.Length > 0
                            && currentPrjMatrixColl.All(prj => prj.UserId != DBMainService.CurrentDBUser.Id))
                {
                    _isProjectCloseToUser = true;
                    MessageBox.Show(
                        $"Вы открыли файл проекта {dBProject.Name}. Данный проект идёт с требованиями от заказчика," +
                        $" и с ними необходимо предварительно ознакомиться. Для этого - обратись в BIM-отдел." +
                        "\nИНФО: Сейчас файл закроется",
                        "KPLN: Ограниченный проект",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);

                    BitrixMessageSender.SendMsg_ToBIMChat(
                        $"Сотрудник: {DBMainService.CurrentDBUser.Surname} {DBMainService.CurrentDBUser.Name} " +
                        $"из отдела {DBMainService.CurrentUserDBSubDepartment.Code}\n" +
                        $"Статус допуска: Данный сотрудник не имеет доступа к этому проекту (его нужно внести в список)\n" +
                        $"Действие: Открыл проект {doc.Title}.");

                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new DocCloser(DBMainService.CurrentDBUser, doc));

                    return;
                }
            }
            else
            {
                TaskDialog td = new TaskDialog("ВНИМАНИЕ")
                {
                    MainIcon = TaskDialogIcon.TaskDialogIconInformation,
                    MainInstruction = "Вы работаете в незарегистрированном проекте. Скинь скрин в BIM-отдел",
                    FooterText = $"Специалисту BIM-отдела: файл - {fileFullName}",
                    CommonButtons = TaskDialogCommonButtons.Ok,
                };
                td.Show();
            }
            #endregion
        }

        /// <summary>
        /// Событие на сохранение файла (НЕ работает при синзронизации, даже если указать о сохранении локалки)
        /// </summary>
        private void OnDocumentSaved(object sender, DocumentSavedEventArgs args)
        {
            Document doc = args.Document;

            ARKonFileSendMsg(doc);
        }

        /// <summary>
        /// Событие на синхронизацию файла
        /// </summary>
        private void OnDocumentSynchronized(object sender, DocumentSynchronizedWithCentralEventArgs args)
        {
            Document doc = args.Document;

            ARKonFileSendMsg(doc);

#if Revit2020 || Revit2023
            if (MonitoredDocFilePath_ExceptARKon(doc) == null)
                return;
#endif

            #region Отлов пользователей с ограничением допуска к работе ВО ВСЕХ ПРОЕКТАХ
            if (DBMainService.CurrentDBUser.IsUserRestricted)
            {
                BitrixMessageSender.SendMsg_ToBIMChat(
                    $"Сотрудник: {DBMainService.CurrentDBUser.Surname} {DBMainService.CurrentDBUser.Name} " +
                    $"из отдела {DBMainService.CurrentUserDBSubDepartment.Code}\n" +
                    $"Статус допуска: Ограничен в работе с реальными проектами (IsUserRestricted={DBMainService.CurrentDBUser.IsUserRestricted})\n" +
                    $"Действие: Произвел синхронизацию файла {doc.Title}.");

                MessageBox.Show(
                    $"Вы произвели синхронизацию проекта с диска Y:\\, хотя у вас нет к этому доступа (вы не сдали КЛ BIM-отделу). " +
                    $"Данные переданы в BIM-отдел и ГИ бюро. Файл будет ЗАКРЫТ.",
                    "KPLN: Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new DocCloser(DBMainService.CurrentDBUser, doc));
            }
            #endregion

            string fileFullName = GetFileFullName(doc);

            #region Обработка работы в архивных копиях
            if (fileFullName.ToLower().Contains("архив"))
            {
                BitrixMessageSender.SendMsg_ToBIMChat(
                    $"Сотрудник: {DBMainService.CurrentDBUser.Surname} {DBMainService.CurrentDBUser.Name} из отдела {DBMainService.CurrentUserDBSubDepartment.Code}\n" +
                    $"Статус допуска: Сотрудник засинхронизировал АРХИВНЫЙ проект\n" +
                    $"Действие: Произвел синхронизацию в {doc.Title}.");

                MessageBox.Show(
                    $"Вы произвели синхронизацию АРХИВНОГО проекта с диска Y:\\. Данные переданы в BIM-отдел. Файл будет ЗАКРЫТ.",
                    "KPLN: Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new DocCloser(DBMainService.CurrentDBUser, doc));
            }
            #endregion

            #region Работа с проектами КПЛН
            // Получаю проект из БД КПЛН
            DBProject dBProject = DBMainService.ProjectDbService.GetDBProject_ByRevitDocFileNameANDRVersion(fileFullName, RevitVersion);
            DBSubDepartment prjDBSubDepartment = DBMainService.SubDepartmentDbService.GetDBSubDepartment_ByRevitDocFullPath(doc.PathName);
            DBDocument dBDocument = DBDocumentByRevitDocPathAndDBProject(fileFullName, dBProject, prjDBSubDepartment);
            if (dBDocument != null)
            {
                DBMainService
                    .DocDbService
                    .UpdateDBDocument_LastChangedData(dBDocument, DBMainService.CurrentDBUser, DBMainService.CurrentTimeForDB());

                // Защита закрытого проекта от изменений (файл вообще не должен открываться, но ЕСЛИ это произошло - будет уведомление)
                if (dBDocument.IsClosed)
                {
                    BitrixMessageSender.SendMsg_ToBIMChat(
                        $"Сотрудник: {DBMainService.CurrentDBUser.Surname} {DBMainService.CurrentDBUser.Name} из отдела {DBMainService.CurrentUserDBSubDepartment.Code}\n" +
                        $"Статус допуска: Сотрудник засинхронизировал ЗАКРЫТЫЙ проект (проект все же НЕ удалось закрыть)\n" +
                        $"Действие: Произвел синхронизацию в {doc.Title}.");

                    MessageBox.Show(
                        $"Вы произвели синхронизацию ЗАКРЫТОГО проекта с диска Y:\\. Данные переданы в BIM-отдел. Файл будет ЗАКРЫТ.",
                        "KPLN: Ошибка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);

                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new DocCloser(DBMainService.CurrentDBUser, doc));
                }
                // Отлов пользователей с ограничением допуска к работе в текущем проекте
                else if (_isProjectCloseToUser)
                {
                    BitrixMessageSender.SendMsg_ToBIMChat(
                        $"Сотрудник: {DBMainService.CurrentDBUser.Surname} {DBMainService.CurrentDBUser.Name} из отдела {DBMainService.CurrentUserDBSubDepartment.Code}\n" +
                        $"Статус допуска: Данный сотрудник не имеет доступа к этому проекту (его нужно внести в список)\n" +
                        $"Действие: Произвел синхронизацию в {doc.Title}.");

                    MessageBox.Show(
                        $"Вы открыли файл проекта {dBProject.Name}. " +
                        $"Данный проект идёт с требованиями от заказчика, и с ними необходимо предварительно ознакомиться. " +
                        $"Для этого - обратись в BIM-отдел. Файл будет ЗАКРЫТ.",
                        "KPLN: Ошибка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);

                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new DocCloser(DBMainService.CurrentDBUser, doc));
                }
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
                    if (doc.PathName.Contains("_КР_") && doc.PathName.Contains("МТРС_С2."))
                        RSBackupFile(doc, "Y:\\Жилые здания\\Матросская Тишина\\10.Стадия_Р\\6.КР\\1.RVT\\00_Автоархив с Revit-Server");
                    else if (doc.PathName.Contains("_ОВ_ТМ_"))
                        RSBackupFile(doc, "Y:\\Жилые здания\\Матросская Тишина\\10.Стадия_Р\\7.4.1.ИТП\\1.RVT\\00_Автоархив с Revit-Server");
                    else if (doc.PathName.Contains("_ОВ_АТМ_"))
                        RSBackupFile(doc, "Y:\\Жилые здания\\Матросская Тишина\\10.Стадия_Р\\7.4.1.ИТП\\1.RVT\\00_Автоархив с Revit-Server");
                }

                // Проект Матросская тишина
                bool isIZML1 = doc.PathName.Contains("ИЗМЛ_");
                if (isIZML1)
                {
                    if (doc.PathName.Contains("_АР_"))
                        RSBackupFile(doc, "Y:\\Жилые здания\\ФСК_Измайловский\\10.Стадия_Р\\5.АР\\1.RVT\\1 очередь\\00_Автоархив с Revit-Server");
                }
            }
            #endregion


            #region Автопроверки (лучше в конец, чтобы всё остальное отработало корректно)
            if (RunAutoChecks)
            {
                UIApplication uiApp = new UIApplication(doc.Application);
                CheckBatchRunner.RunAll(uiApp, doc.Title);
            }
            #endregion
        }
    }
}
