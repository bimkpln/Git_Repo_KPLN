using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Library_Bitrix24Worker;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_SQLiteWorker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Loader.Common;
using KPLN_Looker.Comparers;
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
        /// Кэширование текущего пути мониторингового проекта
        /// </summary>
        private static string _monitoredDocFilePath_ExceptARKon;

        /// <summary>
        /// Кэширование текущего проекта KPLN
        /// </summary>
        private static DBProject _currentDBProject;

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
        private static bool _isProjectCloseToUser;

        /// <summary>
        /// Вариативность аббр. имён разбивочного файла
        /// </summary>
        private static readonly string[] _rfFileNames = new string[]
        {
            "_рф",
            "рф_",
            "-рф",
            "рф-",
            "_разб.файл",
            "разб.файл_",
            "-разб.файл",
            "разб.файл-",
            "_разбфайл",
            "разбфайл_",
            "-разбфайл",
            "разбфайл-",
            "_разбив",
            "разбив_",
            "-разбив",
            "разбив-",
        };

        /// <summary>
        /// Счётчик помещений в модели. Ключ - имя ФХ/модели, значения - коллекция помещений
        /// </summary>
        private static readonly Dictionary<string, Room[]> _lastDocRooms = new Dictionary<string, Room[]>();

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
                application.ControlledApplication.DocumentClosing += OnDocumentClosing;
                application.ViewActivated += OnViewActivated;
                application.ControlledApplication.DocumentChanged += OnDocumentChanged;
                application.ControlledApplication.DocumentSynchronizedWithCentral += OnDocumentSynchronized;

#if DEBUG
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

            string fileFullName = KPLN_Library_SQLiteWorker.FactoryParts.DocumentDbService.GetFileFullName(doc);

            if (!doc.IsFamilyDocument
#if DEBUG
                && (fileFullName.ToLower().Contains("stinproject.local\\project\\") || fileFullName.ToLower().Contains("fs01\\lib\\отдел bim\\") || fileFullName.ToLower().Contains("rsn"))
#else
                && (fileFullName.ToLower().Contains("stinproject.local\\project\\") || fileFullName.ToLower().Contains("rsn"))
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
        /// Событие на открытый документ
        /// </summary>
        private static void OnDocumentOpened(object sender, DocumentOpenedEventArgs args)
        {
            Document doc = args.Document;
            if (doc == null) return;

            #region Утсановка переменных, привязаных к открытому файлу
            // Имя файла
            _monitoredDocFilePath_ExceptARKon = MonitoredDocFilePath_ExceptARKon(doc);


            // Проект КПЛН
            string fileFullName = KPLN_Library_SQLiteWorker.FactoryParts.DocumentDbService.GetFileFullName(doc);
            _currentDBProject = DBMainService.ProjectDbService.GetDBProject_ByRevitDocFileNameANDRVersion(fileFullName, RevitVersion);


            // Кол-во помещений в текущей модели
            Room[] docRooms = GetDocRooms(doc);

            if (_lastDocRooms.ContainsKey(fileFullName))
                _lastDocRooms[fileFullName] = docRooms;
            else
                _lastDocRooms.Add(fileFullName, docRooms);
            #endregion


            // Если проект не мониториться - игнор
            if (_monitoredDocFilePath_ExceptARKon == null)
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
            if (_currentDBProject != null)
            {
                // Ищу документ
                DBSubDepartment prjDBSubDepartment = DBMainService.SubDepartmentDbService.GetDBSubDepartment_ByRevitDocFullPath(fileFullName);
                DBDocument dBDocument = DBDocumentByRevitDocPathAndDBProject(fileFullName, _currentDBProject, prjDBSubDepartment);
                if (dBDocument == null)
                {
                    dBDocument = new DBDocument()
                    {
                        CentralPath = fileFullName,
                        ProjectId = _currentDBProject.Id,
                        SubDepartmentId = prjDBSubDepartment.Id,
                        LastChangedUserId = DBMainService.CurrentDBUser.Id,
                        LastChangedData = DBMainService.CurrentTimeForDB(),
                        IsClosed = false,
                    };

                    DBMainService.DocDbService.CreateDBDocument(dBDocument);
                }

                //Обрабатываю документ
                DBMainService.DocDbService.UpdateDBDocument_IsClosedByProject(_currentDBProject);

                // Вывожу окно, если документ ЗАКРЫТ к редактированию
                if (_currentDBProject.IsClosed)
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

                        BitrixMessageSender.SendMsg_ToUser_ByJSONRequest(jsonRequestToUser);
                    }
                    #endregion

                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new DocCloser(DBMainService.CurrentDBUser, doc));

                    return;
                }


                // Отлов пользователей с ограничением допуска к работе в текущем проекте
                DBProjectsAccessMatrix[] currentPrjMatrixColl = DBMainService
                    .PrjAccessMatrixDbService
                    .GetDBProjectMatrix_ByProject(_currentDBProject)
                    .ToArray();
                if (currentPrjMatrixColl.Length > 0
                    && currentPrjMatrixColl.All(prj => prj.UserId != DBMainService.CurrentDBUser.Id))
                {
                    _isProjectCloseToUser = true;
                    MessageBox.Show(
                        $"Вы открыли файл проекта {_currentDBProject.Name}. Данный проект идёт с требованиями от заказчика," +
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
        /// Событие на закрывающийся документ
        /// </summary>
        private static void OnDocumentClosing(object sender, DocumentClosingEventArgs args)
        {
            Document doc = args.Document;
            if (doc == null) return;

            string fileFullName = KPLN_Library_SQLiteWorker.FactoryParts.DocumentDbService.GetFileFullName(doc);
            if (_lastDocRooms.ContainsKey(fileFullName))
                _lastDocRooms.Remove(fileFullName);
        }

        /// <summary>
        /// Событие на активацию вида
        /// </summary>
        private static void OnViewActivated(object sender, ViewActivatedEventArgs args)
        {
            Document doc = args.Document;
            if (doc == null) return;

            UIDocument uidoc = new UIDocument(doc);
            Autodesk.Revit.DB.View activeView = args.CurrentActiveView;

#if REVIT
            // Игнор НЕ мониторинговых моделей
            if (MonitoredDocFilePath_ExceptARKon(doc) == null)
                return;
#endif
            #region Утсановка переменных, привязаных к виду
            // Имя файла
            _monitoredDocFilePath_ExceptARKon = MonitoredDocFilePath_ExceptARKon(doc);


            // Проект КПЛН
            string fileFullName = KPLN_Library_SQLiteWorker.FactoryParts.DocumentDbService.GetFileFullName(doc);
            _currentDBProject = DBMainService.ProjectDbService.GetDBProject_ByRevitDocFileNameANDRVersion(fileFullName, RevitVersion);
            #endregion

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
                    MainInstruction = "Данный вид предназначен только для bim-отдела. Его запрещено открывать или редактировать, поэтому он закроется",
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
            Document doc = args.GetDocument();
            if (doc == null) return;

#if DEBUG
            // Фильтрация по имени проекта
            if (_monitoredDocFilePath_ExceptARKon != null)
            {
                CheckAndSendError_FamilyInstanceUserHided(args, doc.ActiveView);

                // Анализ загрузки семейств путем копирования
                if (// Отлов проекта ПШМ1.1_РД_ОВ. Делает субчик на нашем компе. Семейства правит сам.
                    !_monitoredDocFilePath_ExceptARKon.Contains("Жилые здания\\Пушкино, Маяковского, 1 очередь\\10.Стадия_Р\\7.4.ОВ\\")
                    // Отлов проекта Школа 825. Его дорабатываем за другой организацией
                    && !_monitoredDocFilePath_ExceptARKon.ToLower().Contains("sh1-"))
                    CheckError_FamilyCopiedFromOtherFile(doc, args);
            }
#else
            // Фильтрация по имени проекта
            if (_monitoredDocFilePath_ExceptARKon != null)
            {
                CheckAndSendError_FamilyInstanceUserHided(args, doc.ActiveView);

                // Анализ загрузки семейств путем копирования
                if (// Отлов проекта ПШМ1.1_РД_ОВ. Делает субчик на нашем компе. Семейства правит сам.
                    !_monitoredDocFilePath_ExceptARKon.Contains("Жилые здания\\Пушкино, Маяковского, 1 очередь\\10.Стадия_Р\\7.4.ОВ\\")
                    // Отлов проекта Школа 825. Его дорабатываем за другой организацией
                    && !_monitoredDocFilePath_ExceptARKon.ToLower().Contains("sh1-"))
                    CheckError_FamilyCopiedFromOtherFile(doc, args);
            }
#endif
        }

        /// <summary>
        /// Событие на синхронизацию файла
        /// </summary>
        private static void OnDocumentSynchronized(object sender, DocumentSynchronizedWithCentralEventArgs args)
        {
            Document doc = args.Document;
            if (doc == null) return;

            ARKonFileSendMsg(doc);

#if REVIT
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

            string fileFullName = KPLN_Library_SQLiteWorker.FactoryParts.DocumentDbService.GetFileFullName(doc);
            
            // Проверка помещений модели
            CheckAndSendError_RoomDeleted(doc, fileFullName);

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
            DBSubDepartment prjDBSubDepartment = DBMainService.SubDepartmentDbService.GetDBSubDepartment_ByRevitDocFullPath(doc.PathName);
            if (_currentDBProject != null && prjDBSubDepartment != null)
            {
                DBDocument dBDocument = DBDocumentByRevitDocPathAndDBProject(fileFullName, _currentDBProject, prjDBSubDepartment);
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
                            $"Вы открыли файл проекта {_currentDBProject.Name}. " +
                            $"Данный проект идёт с требованиями от заказчика, и с ними необходимо предварительно ознакомиться. " +
                            $"Для этого - обратись в BIM-отдел. Файл будет ЗАКРЫТ.",
                            "KPLN: Ошибка",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);

                        KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new DocCloser(DBMainService.CurrentDBUser, doc));
                    }
                }
            }
            #endregion

            #region Бэкап версий с RS на наш сервак по проектам
            if (args.Status == RevitAPIEventStatus.Succeeded && _currentDBProject != null && _currentDBProject.RevitServerPath != null)
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
                    ModelPath mPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(KPLN_Library_SQLiteWorker.FactoryParts.DocumentDbService.GetFileFullName(doc));
                    if (mPath.ServerPath)
                    {
                        if (doc.PathName.Contains("_КР_"))
                            RSBackupFile(doc, "Y:\\Жилые здания\\Матросская Тишина\\10.Стадия_Р\\6.КР\\1.RVT\\00_Автоархив с Revit-Server");
                        else if (doc.PathName.Contains("_ОВ_ТМ_"))
                            RSBackupFile(doc, "Y:\\Жилые здания\\Матросская Тишина\\10.Стадия_Р\\7.4.1.ИТП\\1.RVT\\00_Автоархив с Revit-Server");
                        else if (doc.PathName.Contains("_ОВ_АТМ_"))
                            RSBackupFile(doc, "Y:\\Жилые здания\\Матросская Тишина\\10.Стадия_Р\\7.4.1.ИТП\\1.RVT\\00_Автоархив с Revit-Server");
                    }
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
                // Модели РФ не проверяем вообще (там нет мониторинга, нет линков. РН есть, но они не критичны)
                bool containsRf = _rfFileNames.Any(m => doc.PathName.IndexOf(m, StringComparison.OrdinalIgnoreCase) >= 0);
                if (!containsRf)
                {
                    UIApplication uiApp = new UIApplication(doc.Application);
                    CheckBatchRunner.RunAll(uiApp, doc.Title);
                }
            }
            #endregion
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

            if (_monitoredDocFilePath_ExceptARKon == null
                // Отлов проекта ПШМ1.1_РД_ОВ. Делает субчик на нашем компе. Семейства правит сам.
                || _monitoredDocFilePath_ExceptARKon.Contains("Жилые здания\\Пушкино, Маяковского, 1 очередь\\10.Стадия_Р\\7.4.ОВ\\")
                // Отлов проекта Школа 825. Его дорабатываем за другой организацией
                || _monitoredDocFilePath_ExceptARKon.ToLower().Contains("sh1-"))
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
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                    BuiltInCategory bic = (BuiltInCategory)famCat.Id.IntegerValue;
#else
                    BuiltInCategory bic = (BuiltInCategory)famCat.Id.Value;
#endif

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
        /// Событие на сохранение файла (НЕ работает при синзронизации, даже если указать о сохранении локалки)
        /// </summary>
        private static void OnDocumentSaved(object sender, DocumentSavedEventArgs args)
        {
            Document doc = args.Document;

            ARKonFileSendMsg(doc);
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
        /// Если проверка факта удаления помещения. При наличии - отправка в BIM в битрикс
        /// </summary>
        private static void CheckAndSendError_RoomDeleted(Document doc, string fileFullName)
        {
            // Проверка, что это проект КПЛН стадии РД
            if (_currentDBProject == null || _currentDBProject.Stage != "РД") return;
            
            
            Room[] updatedRoomColl = GetDocRooms(doc);
            if (updatedRoomColl.Length == 0 && _lastDocRooms.ContainsKey(fileFullName) && _lastDocRooms[fileFullName].Length == 0)
                return;

            // Пусть надоедает юзеру, НО не крашит ревит
            if (!_lastDocRooms.ContainsKey(fileFullName))
            {
                HtmlOutput.Print("Ошибка определения модели при анализе на соответсвие помещений. Отправь разработчику!", MessageType.Error);
                return;
            }


            // Старая коллекция ID помещений
            HashSet<ElementId> prevRoomIds = new HashSet<ElementId>(_lastDocRooms[fileFullName].Select(r =>
                {
                    if (r != null && r.IsValidObject)
                        return r.Id;

                    return new ElementId(-1);
                }), ElementIdEqualityComparer.Instance);
            
            
            // Новая коллекция ID помещений
            HashSet<ElementId> currRoomIds = new HashSet<ElementId>(updatedRoomColl.Select(r => r.Id), ElementIdEqualityComparer.Instance);


            // Анализ: Добавлены новые
            IEnumerable<ElementId> addedIds = currRoomIds.Where(id => !prevRoomIds.Contains(id));
            var addedRooms = updatedRoomColl.Where(r => addedIds.Contains(r.Id, ElementIdEqualityComparer.Instance)).ToArray();

            // Анализ: Удаленные
            IEnumerable<ElementId> removedIds = prevRoomIds.Where(id => !currRoomIds.Contains(id));
            var removedRooms = _lastDocRooms[fileFullName].Where(r => removedIds.Contains(r.Id, ElementIdEqualityComparer.Instance)).ToArray();

            // Анализ: Общее число изменений. Если его нет - игнор
            var changedRooms = addedRooms.Concat(removedRooms).ToArray();
            if (changedRooms.Length == 0) return;


            // Обновляем коллекцию кол-ва помещений
            _lastDocRooms[fileFullName] = updatedRoomColl;


            // Проверка, что это модель АР
            DBSubDepartment prjDBSubDepartment = DBMainService.SubDepartmentDbService.GetDBSubDepartment_ByRevitDocFullPath(doc.PathName);
            DBDocument dBDocument = DBDocumentByRevitDocPathAndDBProject(fileFullName, _currentDBProject, prjDBSubDepartment);
            if (dBDocument == null || dBDocument.SubDepartmentId != 2) return;


            // Отправка сообещния в битрикс
            if (DBMainService.CurrentDBUser != null)
            {
                BitrixMessageSender.SendMsg_ToBIMChat(
                    $"Сотрудник: {DBMainService.CurrentDBUser.Surname} {DBMainService.CurrentDBUser.Name} " +
                    $"из отдела {DBMainService.CurrentUserDBSubDepartment.Code}\n" +
                    $"Действие: Произвёл синхронизацию проекта с [b]созданными/удаленными помещениями[/b] в модели {doc.Title}.\n" +
                    $"Инфо: Пересоздание помещений влияет на работу смежных разделов, а также на проверку соответствия ТЭПов между стадией П и Р.\n" +
                    $"Список id помещений из {changedRooms.Count()} шт.: {string.Join(",", changedRooms.Select(el => el.Id))}");
            }
        }

        /// <summary>
        /// Подсчёт помещений в модели
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        private static Room[] GetDocRooms(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Where(el =>
                {
                    if (el is Room room)
                        return room.Area > 0.001;

                    return false;
                })
                .Cast<Room>()
                .ToArray();
        }

        /// <summary>
        /// Поиск ошибки скрытия элементов запрещенными действиями. При наличии - отправка пользователю в битрикс
        /// </summary>
        private static void CheckAndSendError_FamilyInstanceUserHided(DocumentChangedEventArgs args, Autodesk.Revit.DB.View activeView)
        {
            string transName = args.GetTransactionNames().FirstOrDefault();
            if (transName == null
                || activeView == null
                // Игнор чертёжных видов
                || activeView.ViewType == ViewType.DraftingView)
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
        private static void CheckError_FamilyCopiedFromOtherFile(Document doc, DocumentChangedEventArgs args)
        {
            string transName = args.GetTransactionNames().FirstOrDefault();
            if (transName == null)
                return;

            string lowerTransName = transName.ToLower();
            if (!lowerTransName.Contains("встав") && !lowerTransName.Contains("paste"))
                return;

            // Коллекция добавленых анализируемых FamilyInstance
            SetResultFamInstFilterForAddedElems();
            FamilyInstance[] monitoredFamInsts = args
                .GetAddedElementIds(_resultFamInstFilter)
                .Select(id => doc.GetElement(id))
                .Cast<FamilyInstance>()
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                .Where(fi => IsFamilyMonitoredError(doc, (BuiltInCategory)fi.Category.Id.IntegerValue, fi.Symbol.FamilyName))
#else
                .Where(fi => IsFamilyMonitoredError(doc, (BuiltInCategory)fi.Category.Id.Value, fi.Symbol.FamilyName))
#endif
                .ToArray();

            if (!monitoredFamInsts.Any())
                return;

            // Есть возможность копировать листы через буфер обмена. Ревит автоматом добавляет инкременту номеру, имя остаётся тем же
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
            if (monitoredFamInsts.All(fi => fi.Category.Id.IntegerValue == (int)BuiltInCategory.OST_TitleBlocks))
#else
            if (monitoredFamInsts.All(fi => fi.Category.Id.Value == (int)BuiltInCategory.OST_TitleBlocks))
#endif
                return;

            if (_currentDBProject == null)
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
                string openDocFileFullName = KPLN_Library_SQLiteWorker.FactoryParts.DocumentDbService.GetFileFullName(openDoc);
                DBProject openDocDBProject = DBMainService.ProjectDbService.GetDBProject_ByRevitDocFileNameANDRVersion(openDocFileFullName, RevitVersion);
                if (openDocDBProject != null && _currentDBProject.Id == openDocDBProject.Id)
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

            string fileFullName = KPLN_Library_SQLiteWorker.FactoryParts.DocumentDbService.GetFileFullName(doc);
            // Проекта нет, а путь принадлежит к мониторинговым ВКЛЮЧАЯ концепции - то оповещение о новом проекте
            if (_currentDBProject == null)
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
            DBDocument dBDocument = DBDocumentByRevitDocPathAndDBProject(fileFullName, _currentDBProject, prjDBSubDepartment);
            if (dBDocument == null)
            {
                foreach (DBUser dbUser in _arKonFileSubscribersFromBIM)
                {
                    BitrixMessageSender.SendMsg_ToUser_ByDBUser(
                        dbUser,
                        $"Сотрудник: {DBMainService.CurrentDBUser.Surname} {DBMainService.CurrentDBUser.Name} " +
                        $"из отдела {DBMainService.CurrentUserDBSubDepartment.Code}\n" +
                        $"Действие: Произвел сохранение/синхронизацию нового файла проекта [b]{_currentDBProject.Name}_{_currentDBProject.Stage}[/b] (сообщение возникает только при 1м сохранении).\n" +
                        $"Имя файла: [b]{doc.Title}[/b].\n" +
                        $"Путь к модели: [b]{fileFullName}[/b].",
                        "Y");
                }

                // Создаю, если не нашел
                dBDocument = new DBDocument()
                {
                    CentralPath = fileFullName,
                    ProjectId = _currentDBProject.Id,
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
    }
}
