using Autodesk.Revit.DB;
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
using System.Text.RegularExpressions;
using System.Windows.Forms;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Looker
{
    public class Module : IExternalModule
    {
        private static int _restrictedUserSynch = 0;
        private readonly DBWorkerService _dBWorkerService;
        private bool _isProjectCloseToUser = false;

        public Module()
        {
            _dBWorkerService = new DBWorkerService();
        }

        /// <summary>
        /// Указатель на окно ревит
        /// </summary>
        public static IntPtr MainWindowHandle { get; set; }

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
                INIFileService iNIFileService = new INIFileService(_dBWorkerService.CurrentDBUser, application.ControlledApplication.VersionNumber);
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
                if (!_dBWorkerService.CurrentDBUserSubDepartment.Code.ToUpper().Contains("BIM"))
                    application.ControlledApplication.FamilyLoadingIntoDocument += OnFamilyLoadingIntoDocument;

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
        private void OnApplicationInitialized(object sender, ApplicationInitializedEventArgs e)
        {
            if (_dBWorkerService.CurrentDBUser.IsUserRestricted)
            {
                MessageBox.Show(
                    $"Ваша работа ограничена работой в тестовых файлах. Любой факт попытки открытия/синхронизации при работе с файлами с диска Y:\\ - будет передан в BIM-отдел",
                    "Предупреждение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Asterisk);
            }
        }

        /// <summary>
        /// Событие на открытие документа
        /// </summary>
        private void OnDocumentOpened(object sender, DocumentOpenedEventArgs args)
        {
            Document doc = args.Document;
            if (MonitoredFilePath(doc) != null)
            {
                string centralPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath());
                #region Отлов пользователей с ограничением допуска к работе ВО ВСЕХ ПРОЕКТАХ
                if (_dBWorkerService.CurrentDBUser.IsUserRestricted)
                {
                    BitrixMessageSender.SendMsg_ToBIMChat(
                        $"Сотрудник: {_dBWorkerService.CurrentDBUser.Surname} {_dBWorkerService.CurrentDBUser.Name} из отдела {_dBWorkerService.CurrentDBUserSubDepartment.Code}\n" +
                        $"Статус допуска: Ограничен в работе с реальными проектами (IsUserRestricted={_dBWorkerService.CurrentDBUser.IsUserRestricted})\n" +
                        $"Действие: Открыл проект {doc.Title}.");

                    MessageBox.Show(
                        $"Вы открытли проект с диска Y:\\. Напомню - Ваша работа ограничена тестовыми файлами! Данные переданы в BIM-отдел",
                        "Ошибка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
                #endregion

                #region Обработка проектов КПЛН
                DBProject dBProject = _dBWorkerService.Get_DBProjectByRevitDocFile(centralPath);
                // У Сетуни 2 ревит сервера, что являеться жестким исключением, поэтому её захардкодил сюда
                if (centralPath.Contains("Самолет_Сетунь") && dBProject == null)
                {
                    string[] splitName = centralPath.Split(new string[] { "RSN://rs01/" }, StringSplitOptions.None);
                    centralPath = Path.Combine("RSN://192.168.0.5/", splitName[1]);
                    dBProject = _dBWorkerService.Get_DBProjectByRevitDocFile(centralPath);
                }

                if (dBProject != null)
                {
                    // Ищу документ
                    DBDocument dBDocument = _dBWorkerService.Get_DBDocumentByRevitDocPathAndDBProject(centralPath, dBProject);
                    if (dBDocument == null)
                    {
                        // Создаю, если не нашел
                        DBSubDepartment dBSubDepartment = _dBWorkerService.CurrentDBUserSubDepartment;
                        DBUser dbUser = _dBWorkerService.CurrentDBUser;
                        dBDocument = _dBWorkerService.Create_DBDocument(
                            centralPath,
                            dBProject.Id,
                            dBSubDepartment.Id,
                            dbUser.Id,
                            _dBWorkerService.CurrentTimeForDB(),
                            false);
                    }

                    //Обрабатываю докемент
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
                        _dBWorkerService.Update_DBDocumentIsClosedStatus(dBProject);
                        DBProjectMatrix[] currentPrjMatrixColl = _dBWorkerService.CurrentDBProjectMatrixColl.Where(prj => dBProject.Id == prj.ProjectId).ToArray();
                        // Вывожу окно, если документ ЗАКРЫТ к редактированию
                        if (dBProject != null && dBProject.IsClosed)
                        {
                            TaskDialog taskDialog = new TaskDialog("KPLN: Закрытый проект")
                            {
                                MainIcon = TaskDialogIcon.TaskDialogIconError,
                                MainContent = "Вы пытаетесь работать в закрытом проекте. О факте синхранизации узнает BIM-отдел. " +
                                    "Чтобы получить доступ на внесение изменений в этот проект - обратитесь в BIM-отдел",
                                CommonButtons = TaskDialogCommonButtons.Ok,
                            };
                            taskDialog.Show();
                        }
                        // Отлов пользователей с ограничением допуска к работе в текщем проекте
                        else if (currentPrjMatrixColl.Length > 0 && !currentPrjMatrixColl.Where(prj => prj.UserId == _dBWorkerService.CurrentDBUser.Id).Any())
                        {
                            _isProjectCloseToUser = true;
                            BitrixMessageSender.SendMsg_ToBIMChat(
                                $"Сотрудник: {_dBWorkerService.CurrentDBUser.Surname} {_dBWorkerService.CurrentDBUser.Name} из отдела {_dBWorkerService.CurrentDBUserSubDepartment.Code}\n" +
                                $"Статус допуска: Данный сотрудник не имеет доступа к этому проекту (его нужно внести в список)\n" +
                                $"Действие: Открыл проект {doc.Title}.");

                            MessageBox.Show(
                                $"Вы открытли файл проекта {dBProject.Name}. Данный проект идёт с требованиями от заказчика, и с ними необходимо предварительно ознакомиться. Для этого - обратись в BIM-отдел." +
                                $"\nИНФО: Если произвести синзронизацию проекта - он закроется",
                                "Ошибка",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
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

        }

        /// <summary>
        /// Событие на активацию вида
        /// </summary>
        private void OnViewActivated(object sender, ViewActivatedEventArgs args)
        {
            Autodesk.Revit.DB.View activeView = args.CurrentActiveView;
            #region Закрываю вид, если он для бим-отдела
            if (MonitoredFilePath(args.Document) != null)
            {
                if (activeView != null
                    && activeView is View3D _
                    && (activeView.Title.ToUpper().Contains("BIM360")
                        || activeView.Title.ToUpper().Contains("NAVISWORKS")
                        || activeView.Title.ToUpper().Contains("GSTATION")
                        || activeView.Title.ToUpper().Contains("NWC")
                        || activeView.Title.ToUpper().Contains("NWD"))
                    && !_dBWorkerService.CurrentDBUserSubDepartment.Code.ToUpper().Contains("BIM"))
                {
                    TaskDialog td = new TaskDialog("ВНИМАНИЕ")
                    {
                        MainIcon = TaskDialogIcon.TaskDialogIconError,
                        MainInstruction = "Данный вид предназначен только для bim-отдела. Его запрещено открывать или редактировать. Вид будет закрыт",
                        CommonButtons = TaskDialogCommonButtons.Ok,
                    };
                    td.Show();

                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new ViewCloser(activeView.Id));
                }
            }
            #endregion
        }

        /// <summary>
        /// Событие на изменение в документе
        /// </summary>
        private void OnDocumentChanged(object sender, DocumentChangedEventArgs args)
        {
            #region Анализ загрузки семейств путем копирования
            Document doc = args.GetDocument();
            if (MonitoredFilePath(doc) != null)
                IsFamilyLoadedFromOtherFile(args);
            #endregion
        }

        /// <summary>
        /// Событие на синхронизацию файла
        /// </summary>
        private void OnDocumentSynchronized(object sender, DocumentSynchronizedWithCentralEventArgs args)
        {
            Document doc = args.Document;
            if (MonitoredFilePath(doc) != null)
            {
                string centralPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath());
                #region Отлов пользователей с ограничением допуска к работе ВО ВСЕХ ПРОЕКТАХ
                if (_dBWorkerService.CurrentDBUser.IsUserRestricted)
                {
                    _restrictedUserSynch++;
                    BitrixMessageSender.SendMsg_ToBIMChat(
                        $"Сотрудник: {_dBWorkerService.CurrentDBUser.Surname} {_dBWorkerService.CurrentDBUser.Name} из отдела {_dBWorkerService.CurrentDBUserSubDepartment.Code}\n" +
                        $"Статус допуска: Ограничен в работе с реальными проектами (IsUserRestricted={_dBWorkerService.CurrentDBUser.IsUserRestricted})\n" +
                        $"Действие: Произвел синхронизацию проекта {doc.Title}.");

                    if (_restrictedUserSynch == 1)
                        MessageBox.Show(
                            $"Вы произвели синхронизацию проекта с диска Y:\\, хотя у вас нет к этому доступа (вы не сдали КЛ BIM-отделу). Данные переданы в BIM-отдел и ГИ бюро. Следующая синхнронизация приведёт к ЗАКРЫТИЮ вашего файла.",
                            "Ошибка",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    else
                    {
                        MessageBox.Show(
                            $"Вы произвели синхронизацию проекта с диска Y:\\, хотя у вас нет к этому доступа (вы не сдали КЛ BIM-отделу). Данные переданы в BIM-отдел и ГИ бюро. Файл будет ЗАКРЫТ.",
                            "Ошибка",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);

                        KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new DocCloser());
                    }
                }
                #endregion

                #region Работа с проектами КПЛН
                DBProject dBProject = _dBWorkerService.Get_DBProjectByRevitDocFile(centralPath);
                if (dBProject != null)
                {
                    DBDocument dBDocument = _dBWorkerService.Get_DBDocumentByRevitDocPathAndDBProject(centralPath, dBProject);
                    if (dBDocument != null)
                    {
                        _dBWorkerService.Update_DBDocumentLastChangedData(dBDocument);
                        // Защита закрытого проекта от изменений
                        if (dBDocument.IsClosed)
                        {
                            BitrixMessageSender
                                .SendMsg_ToBIMChat($"Сотрудник: {_dBWorkerService.CurrentDBUser.Surname} {_dBWorkerService.CurrentDBUser.Name} из отдела {_dBWorkerService.CurrentDBUserSubDepartment.Code}\n" +
                                $"Действие: Синхронизиция в проекте {doc.Title}, который ЗАКРЫТ.");

                            MessageBox.Show(
                                $"Вы произвели синхронизацию ЗАКРЫТОГО проекта с диска Y:\\. Данные переданы в BIM-отдел.",
                                "Ошибка",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                        }
                        // Отлов пользователей с ограничением допуска к работе в текщем проекте
                        else if (_isProjectCloseToUser)
                        {
                            BitrixMessageSender.SendMsg_ToBIMChat(
                                $"Сотрудник: {_dBWorkerService.CurrentDBUser.Surname} {_dBWorkerService.CurrentDBUser.Name} из отдела {_dBWorkerService.CurrentDBUserSubDepartment.Code}\n" +
                                $"Статус допуска: Данный сотрудник не имеет доступа к этому проекту (его нужно внести в список)\n" +
                                $"Действие: Открыл проект {doc.Title}.");

                            MessageBox.Show(
                                $"Вы открытли файл проекта {dBProject.Name}. Данный проект идёт с требованиями от заказчика, и с ними необходимо предварительно ознакомиться. Для этого - обратись в BIM-отдел. Файл будет ЗАКРЫТ.",
                                "Ошибка",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);

                            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new DocCloser());
                        }
                    }
                }
                #endregion
            }
        }

        /// <summary>
        /// Контроль процесса загрузки семейств в проекты КПЛН
        /// </summary>
        private void OnFamilyLoadingIntoDocument(object sender, FamilyLoadingIntoDocumentEventArgs args)
        {
            Autodesk.Revit.ApplicationServices.Application app = sender as Autodesk.Revit.ApplicationServices.Application;
            Document prjDoc = args.Document;
            string familyName = args.FamilyName;
            string familyPath = args.FamilyPath;

            string docPath = MonitoredFilePath(prjDoc);
            if (docPath == null)
                return;

            #region Игнорирую семейства, которые могут редактировать проектировщики
            DocumentSet appDocsSet = app.Documents;
            foreach (Document doc in appDocsSet)
            {
                if (doc.Title.Contains($"{familyName}"))
                {
                    if (doc.IsFamilyDocument)
                    {
                        Family family = doc.OwnerFamily;
                        Category famCat = family.FamilyCategory;
                        BuiltInCategory bic = (BuiltInCategory)famCat.Id.IntegerValue;

                        // Отлов семейств марок (могут разрабатывать все)
                        if (bic.Equals(BuiltInCategory.OST_ProfileFamilies)
                            || bic.Equals(BuiltInCategory.OST_DetailComponents)
                            || bic.Equals(BuiltInCategory.OST_GenericAnnotation)
                            || bic.Equals(BuiltInCategory.OST_DetailComponentsHiddenLines)
                            || bic.Equals(BuiltInCategory.OST_DetailComponentTags))
                            return;

                        // Отлов семейств марок (могут разрабатывать все), за исключением штампов, подписей и жуков
                        if (famCat.CategoryType.Equals(CategoryType.Annotation)
                            && !familyName.StartsWith("020_")
                            && !familyName.StartsWith("022_")
                            && !familyName.StartsWith("023_")
                            && !familyName.ToLower().Contains("жук"))
                            return;

                        // Отлов семейств лестничных маршей и площадок, которые по форме зависят от проектов (могут разрабатывать все)
                        if (bic.Equals(BuiltInCategory.OST_GenericModel)
                            && (familyName.StartsWith("208_") || familyName.StartsWith("209_")))
                            return;

                        // Отлов семейств соед. деталей каб. лотков производителей: Ostec, Dkc
                        if (bic.Equals(BuiltInCategory.OST_CableTrayFitting)
                            && (familyName.ToLower().Contains("ostec") || familyName.ToLower().Contains("dkc")))
                            return;
                    }
                    else
                        throw new Exception("Ошибка определения типа файла. Обратись к разработчику!");
                }
            }
            #endregion

            #region Отлов семейств, расположенных не на Х, не из плагинов и не из исключений выше (ФИНАЛИЗАЦИЯ ОШИБКИ И ВЫВОД ОКНА ПОЛЬЗОВАТЕЛЮ)
            if (!familyPath.StartsWith("X:\\BIM")
                && !familyPath.Contains("KPLN_Loader")
                // Отлов проекта Школа 825. Его дорабатываем за другой организацией
                && !docPath.ToLower().Contains("sh1-"))
            {
                UserVerify userVerify = new UserVerify("[BEP]: Загружать семейства можно только с диска X");
                userVerify.ShowDialog();

                if (userVerify.Status == UIStatus.RunStatus.CloseBecauseError)
                {
                    TaskDialog.Show("Заперщено", "Не верный пароль, в загрузке семейства отказано!");
                    args.Cancel();
                }
                else if (userVerify.Status == UIStatus.RunStatus.Close)
                {
                    args.Cancel();
                }
            }
            #endregion

        }

        /// <summary>
        /// Выдача имени файла с проверкой на необходимость в контроле действий
        /// </summary>
        private string MonitoredFilePath(Document doc)
        {
            string fileName;
            if (doc.IsWorkshared)
                fileName = ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath());
            else
                fileName = doc.PathName;

            if (doc.IsWorkshared
                && !doc.IsDetached
                && !doc.IsFamilyDocument
                && (fileName.ToLower().Contains("stinproject.local\\project\\") || fileName.ToLower().Contains("rsn"))
                && !fileName.EndsWith("rte")
                && !fileName.ToLower().Contains("\\lib\\")
                && !fileName.ToLower().Contains("конц")
                && !fileName.ToLower().Contains("kon"))
            {
                return fileName;
            }

            return null;
        }

        /// <summary>
        /// Анализ семейств, загруженных из другого проекта (ctrl+c/ctrl+v)
        /// </summary>
        private void IsFamilyLoadedFromOtherFile(DocumentChangedEventArgs args)
        {
            Document doc = args.GetDocument();
            string transName = args.GetTransactionNames().FirstOrDefault();
            if (transName.Contains("Начальная вставка"))
            {
                List<FamilySymbol> addedFamilySymbols = new List<FamilySymbol>();
                ICollection<ElementId> addedElems = args.GetAddedElementIds();
                if (addedElems.Count() > 0)
                {
                    foreach (ElementId elemId in addedElems)
                    {
                        if (doc.GetElement(elemId) is FamilySymbol familySymbol)
                            addedFamilySymbols.Add(familySymbol);
                    }
                }

                if (addedFamilySymbols.Count() > 0)
                {
                    Element[] prjFamilies = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilySymbol))
                        .WhereElementIsElementType()
                        .Where(fs => fs.Category.CategoryType == CategoryType.Model)
                        .ToArray();
                    bool isFamilyInclude = false;
                    bool isFamilyNew = false;
                    foreach (FamilySymbol fs in addedFamilySymbols)
                    {
                        string fsName = fs.FamilyName;
                        string digitEndTrimmer = Regex.Match(fsName, @"\d*$").Value;
                        // Осуществляю срез имени на найденные цифры в конце имени
                        string truePartOfName = fsName.TrimEnd(Regex.Match(fsName, @"\d*$").Value.ToArray());
                        var includeFam = prjFamilies
                            .FirstOrDefault(f => f.Name.Equals(fsName.TrimEnd(Regex.Match(fsName, @"\d*$").Value.ToArray())) && !f.Name.Equals(fsName));

                        if (includeFam == null)
                            isFamilyNew = true;
                        else
                            isFamilyInclude = true;
                    }

                    if (isFamilyInclude && isFamilyNew)
                        MessageBox.Show(
                            "Только что были скопированы семейства, которые являются как новыми, так и уже имеющимися в проекте. " +
                            "Запусти плагин \"KPLN: Проверка семейств\" для проверки семейств на первоисточник",
                            "Предупреждение",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Asterisk);
                    else if (isFamilyInclude)
                        MessageBox.Show(
                            "Только что были скопированы семейства, которые уже имеющимися в проекте. " +
                            "Запусти плагин \"KPLN: Проверка семейств\" для проверки семейств, чтобы избежать дублирования семейств",
                            "Предупреждение",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Asterisk);
                    else if (isFamilyNew)
                        MessageBox.Show(
                            "Только что были скопированы семейства, которые являются новыми. " +
                            "Запусти плагин \"KPLN: Проверка семейств\" для проверки семейств, чтобы избежать наличия семейств из сторонних источников",
                            "Предупреждение",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Asterisk);
                }
            }
        }
    }
}
