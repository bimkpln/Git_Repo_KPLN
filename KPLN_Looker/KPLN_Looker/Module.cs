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
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Looker
{
    public class Module : IExternalModule
    {
        private bool _isProjectCloseToUser = false;

        public Module()
        {
            DBWorkerService = new DBWorkerService();
        }
        public static DBWorkerService DBWorkerService { get; private set; }

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
                INIFileService iNIFileService = new INIFileService(DBWorkerService.CurrentDBUser, application.ControlledApplication.VersionNumber);
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
#if DEBUG
                application.ControlledApplication.FamilyLoadingIntoDocument += OnFamilyLoadingIntoDocument;
#else
                if (!DBWorkerService.CurrentDBUserSubDepartment.Code.ToUpper().Contains("BIM"))
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
        /// Событие, которое будет выполнено при инициализации приложения
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnApplicationInitialized(object sender, ApplicationInitializedEventArgs e)
        {
            if (DBWorkerService.CurrentDBUser.IsUserRestricted)
            {
                MessageBox.Show(
                    $"Ваша работа ограничена работой в тестовых файлах. Любой факт попытки открытия/синхронизации при работе с файлами с диска Y:\\ - будет передан в BIM-отдел",
                    "Предупреждение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Asterisk);
            }
        }

        /// <summary>
        /// Событие на открытый документ
        /// </summary>
        private void OnDocumentOpened(object sender, DocumentOpenedEventArgs args)
        {
            Document doc = args.Document;
            if (MonitoredFilePath(doc) != null)
            {
                string centralPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath());
                #region Отлов пользователей с ограничением допуска к работе ВО ВСЕХ ПРОЕКТАХ
                if (DBWorkerService.CurrentDBUser.IsUserRestricted)
                {
                    BitrixMessageSender.SendMsg_ToBIMChat(
                        $"Сотрудник: {DBWorkerService.CurrentDBUser.Surname} {DBWorkerService.CurrentDBUser.Name} из отдела {DBWorkerService.CurrentDBUserSubDepartment.Code}\n" +
                        $"Статус допуска: Ограничен в работе с реальными проектами (IsUserRestricted={DBWorkerService.CurrentDBUser.IsUserRestricted})\n" +
                        $"Действие: Открыл файл {doc.Title}.");

                    MessageBox.Show(
                        $"Вы открытли проект с диска Y:\\. Напомню - Ваша работа ограничена тестовыми файлами! Данные переданы в BIM-отдел",
                        "KPLN: Ошибка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
                #endregion

                #region Обработка проектов КПЛН
                DBProject dBProject = DBWorkerService.Get_DBProjectByRevitDocFile(centralPath);
                // У Сетуни 2 ревит сервера, что являеться жестким исключением, поэтому её захардкодил сюда
                if (centralPath.Contains("Самолет_Сетунь") && dBProject == null)
                {
                    string[] splitName = centralPath.Split(new string[] { "RSN://rs01/" }, StringSplitOptions.None);
                    centralPath = Path.Combine("RSN://192.168.0.5/", splitName[1]);
                    dBProject = DBWorkerService.Get_DBProjectByRevitDocFile(centralPath);
                }

                if (dBProject != null)
                {
                    // Ищу документ
                    DBDocument dBDocument = DBWorkerService.Get_DBDocumentByRevitDocPathAndDBProject(centralPath, dBProject);
                    if (dBDocument == null)
                    {
                        // Создаю, если не нашел
                        DBSubDepartment dBSubDepartment = DBWorkerService.CurrentDBUserSubDepartment;
                        DBUser dbUser = DBWorkerService.CurrentDBUser;
                        dBDocument = DBWorkerService.Create_DBDocument(
                            centralPath,
                            dBProject.Id,
                            dBSubDepartment.Id,
                            dbUser.Id,
                            DBWorkerService.CurrentTimeForDB(),
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
                        DBWorkerService.Update_DBDocumentIsClosedStatus(dBProject);
                        DBProjectMatrix[] currentPrjMatrixColl = DBWorkerService.CurrentDBProjectMatrixColl.Where(prj => dBProject.Id == prj.ProjectId).ToArray();
                        // Вывожу окно, если документ ЗАКРЫТ к редактированию
                        if (dBProject != null && dBProject.IsClosed)
                        {
                            MessageBox.Show(
                                "Вы открыли ЗАКРЫТЫЙ проект. Работа в нём запрещена!\nЧтобы получить доступ на внесение изменений в этот проект - обратитесь в BIM-отдел" +
                                "\nИНФО: Сейчас файл закроется",
                                "KPLN: Закрытый проект",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);

                            BitrixMessageSender.SendMsg_ToBIMChat(
                                $"Сотрудник: {DBWorkerService.CurrentDBUser.Surname} {DBWorkerService.CurrentDBUser.Name} из отдела {DBWorkerService.CurrentDBUserSubDepartment.Code}\n" +
                                $"Статус допуска: Сотрудник открыл ЗАКРЫТЫЙ проект\n" +
                                $"Действие: Открыл файл {doc.Title}.");

                            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new DocCloser(DBWorkerService.CurrentDBUser, doc));
                        }
                        // Отлов пользователей с ограничением допуска к работе в текщем проекте
                        else if (currentPrjMatrixColl.Length > 0 && !currentPrjMatrixColl.Where(prj => prj.UserId == DBWorkerService.CurrentDBUser.Id).Any())
                        {
                            _isProjectCloseToUser = true;
                            MessageBox.Show(
                                $"Вы открытли файл проекта {dBProject.Name}. Данный проект идёт с требованиями от заказчика, и с ними необходимо предварительно ознакомиться. Для этого - обратись в BIM-отдел." +
                                $"\nИНФО: Если файл засинхронизировать - он закроется",
                                "KPLN: Ограниченный проект",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);

                            BitrixMessageSender.SendMsg_ToBIMChat(
                                $"Сотрудник: {DBWorkerService.CurrentDBUser.Surname} {DBWorkerService.CurrentDBUser.Name} из отдела {DBWorkerService.CurrentDBUserSubDepartment.Code}\n" +
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
                    && !DBWorkerService.CurrentDBUserSubDepartment.Code.ToUpper().Contains("BIM"))
                {
                    Document doc = args.Document;
                    UIDocument uidoc = new UIDocument(doc);
                    if (uidoc != null)
                    {
                        IList<UIView> openViews = uidoc.GetOpenUIViews();

                        TaskDialog td;
                        if (openViews.Count > 1)
                        {
                            td = new TaskDialog("ВНИМАНИЕ")
                            {
                                MainIcon = TaskDialogIcon.TaskDialogIconError,
                                MainInstruction = "Данный вид предназначен только для bim-отдела. Его запрещено открывать или редактировать, поэтому он зароектся",
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
                    }
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
            
            #region Бэкап версий с РС на наш сервак по проекту Сетунь
            if (args.Status == RevitAPIEventStatus.Succeeded)
            {
                // Хардкод для старой версии - бэкапим только проект Сетунь
                if (doc.PathName.Contains("СЕТ_1") && doc.PathName.Contains("_АР_"))
                    RSBackupFile(doc, "Y:\\Жилые здания\\project\\Самолет Сетунь\\10.Стадия_Р\\5.АР\\1.RVT\\00_Автоархив с Revit-Server");
                else if (doc.PathName.Contains("СЕТ_1") && doc.PathName.Contains("_КР_"))
                    RSBackupFile(doc, "Y:\\Жилые здания\\Самолет Сетунь\\10.Стадия_Р\\6.КР\\1.RVT\\00_Автоархив с Revit-Server");
                else if (doc.PathName.Contains("СЕТ_1") && doc.PathName.Contains("_ЭОМ"))
                    RSBackupFile(doc, "Y:\\Жилые здания\\Самолет Сетунь\\10.Стадия_Р\\7.1.ЭОМ\\1.RVT\\00_Автоархив с Revit-Server");
                else if (doc.PathName.Contains("СЕТ_1") && doc.PathName.Contains("_ВК"))
                    RSBackupFile(doc, "Y:\\Жилые здания\\Самолет Сетунь\\10.Стадия_Р\\7.2.ВК\\1.RVT\\00_Автоархив с Revit-Server");
                else if (doc.PathName.Contains("СЕТ_1") && doc.PathName.Contains("_ПТ"))
                    RSBackupFile(doc, "Y:\\Жилые здания\\Самолет Сетунь\\10.Стадия_Р\\7.3.АУПТ\\1.RVT\\00_Автоархив с Revit-Server");
                else if (doc.PathName.Contains("СЕТ_1") && doc.PathName.Contains("_ОВ"))
                    RSBackupFile(doc, "Y:\\Жилые здания\\Самолет Сетунь\\10.Стадия_Р\\7.4.ОВ\\1.RVT\\00_Автоархив с Revit-Server");
                else if (doc.PathName.Contains("СЕТ_1") && (doc.PathName.Contains("_ПБ_") || doc.PathName.Contains("_АК_") || doc.PathName.Contains("_СС_")))
                    RSBackupFile(doc, "Y:\\Жилые здания\\Самолет Сетунь\\10.Стадия_Р\\7.5.СС\\1.RVT\\00_Автоархив с Revit-Server");
            }
            #endregion

            if (MonitoredFilePath(doc) != null)
            {
                string centralPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath());
                #region Отлов пользователей с ограничением допуска к работе ВО ВСЕХ ПРОЕКТАХ
                if (DBWorkerService.CurrentDBUser.IsUserRestricted)
                {
                    BitrixMessageSender.SendMsg_ToBIMChat(
                        $"Сотрудник: {DBWorkerService.CurrentDBUser.Surname} {DBWorkerService.CurrentDBUser.Name} из отдела {DBWorkerService.CurrentDBUserSubDepartment.Code}\n" +
                        $"Статус допуска: Ограничен в работе с реальными проектами (IsUserRestricted={DBWorkerService.CurrentDBUser.IsUserRestricted})\n" +
                        $"Действие: Произвел синхронизацию файла {doc.Title}.");
                    
                    MessageBox.Show(
                        $"Вы произвели синхронизацию проекта с диска Y:\\, хотя у вас нет к этому доступа (вы не сдали КЛ BIM-отделу). Данные переданы в BIM-отдел и ГИ бюро. Файл будет ЗАКРЫТ.",
                        "KPLN: Ошибка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);

                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new DocCloser(DBWorkerService.CurrentDBUser, doc));
                }
                #endregion

                #region Работа с проектами КПЛН
                DBProject dBProject = DBWorkerService.Get_DBProjectByRevitDocFile(centralPath);
                if (dBProject != null)
                {
                    DBDocument dBDocument = DBWorkerService.Get_DBDocumentByRevitDocPathAndDBProject(centralPath, dBProject);
                    if (dBDocument != null)
                    {
                        DBWorkerService.Update_DBDocumentLastChangedData(dBDocument);
                        // Защита закрытого проекта от изменений (файл вообще не должен открываться, но ЕСЛИ это произошло - будет уведомление)
                        if (dBDocument.IsClosed)
                        {
                            BitrixMessageSender.SendMsg_ToBIMChat(
                                $"Сотрудник: {DBWorkerService.CurrentDBUser.Surname} {DBWorkerService.CurrentDBUser.Name} из отдела {DBWorkerService.CurrentDBUserSubDepartment.Code}\n" +
                                $"Статус допуска: Сотрудник засинхронизировал ЗАКРЫТЫЙ проект (проект все же НЕ удалось закрыть)\n" +
                                $"Действие: Произвел синхронизацию в {doc.Title}.");

                            MessageBox.Show(
                                $"Вы произвели синхронизацию ЗАКРЫТОГО проекта с диска Y:\\. Данные переданы в BIM-отдел. Файл будет ЗАКРЫТ.",
                                "KPLN: Ошибка",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);

                            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new DocCloser(DBWorkerService.CurrentDBUser, doc));
                        }
                        // Отлов пользователей с ограничением допуска к работе в текщем проекте
                        else if (_isProjectCloseToUser)
                        {
                            BitrixMessageSender.SendMsg_ToBIMChat(
                                $"Сотрудник: {DBWorkerService.CurrentDBUser.Surname} {DBWorkerService.CurrentDBUser.Name} из отдела {DBWorkerService.CurrentDBUserSubDepartment.Code}\n" +
                                $"Статус допуска: Данный сотрудник не имеет доступа к этому проекту (его нужно внести в список)\n" +
                                $"Действие: Произвел синхронизацию в {doc.Title}.");

                            MessageBox.Show(
                                $"Вы открытли файл проекта {dBProject.Name}. Данный проект идёт с требованиями от заказчика, и с ними необходимо предварительно ознакомиться. Для этого - обратись в BIM-отдел. Файл будет ЗАКРЫТ.",
                                "KPLN: Ошибка",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);

                            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new DocCloser(DBWorkerService.CurrentDBUser, doc));
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

                        // Отлов семейств ферм, которые по форме зависят от проектов (могут разрабатывать КР)
                        if ((DBWorkerService.CurrentDBUserSubDepartment.Code.ToUpper().Contains("BIM")
                            || DBWorkerService.CurrentDBUserSubDepartment.Code.ToUpper().Contains("КР"))
                            && bic.Equals(BuiltInCategory.OST_Truss))
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

        /// <summary>
        /// У РС нет возможности откатиться. Делаю свой бэкап на наш сервак
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="pathTo"></param>
        /// <returns></returns>
        private bool RSBackupFile(Document doc, string pathTo)
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
                int archFilesLimit = 10;
                string[] namePrepareArr = doc.Title.Split(new[] { $"_{DBWorkerService.CurrentDBUser.SystemName}" }, StringSplitOptions.None);
                if (namePrepareArr.Length == 0)
                    clearTaskMsg = $"Не удалось определить имя модели из хранилища для пользователя {DBWorkerService.CurrentDBUser.SystemName}. Обратись в BIM-отдел!";
                else
                {
                    string centralFileName = namePrepareArr[0];
                    if (Directory.Exists(pathTo))
                    {
                        string[] archFiles = Directory.GetFiles(pathTo);
                        FileInfo[] currentCentralArchCopies = archFiles
                            .Where(a => a.Contains(centralFileName))
                            .Select(a => new FileInfo(a))
                            .OrderBy(fi => fi.CreationTime)
                            .ToArray();

                        if (currentCentralArchCopies.Count() > archFilesLimit)
                        {
                            int startCount = currentCentralArchCopies.Count() - archFilesLimit;
                            while (startCount > 0)
                            {
                                startCount--;
                                FileInfo archCopyToDel = currentCentralArchCopies[startCount];
                                archCopyToDel.Delete();
                            }
                        }
                    }
                }
            });
            Task.WaitAll(new Task[2] { copyLocalFileTask, clearArchCopyFilesTask });

            if (copyTaskMsg != string.Empty)
                Print($"Ошибка при копировании резервного файла: {copyTaskMsg}", MessageType.Error);

            if (clearTaskMsg != string.Empty)
                Print($"Ошибка при очистке старых резервных копий: {clearTaskMsg}", MessageType.Error);
            return false;
        }

    }
}
