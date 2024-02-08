using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Library_Bitrix24Worker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Loader.Common;
using KPLN_Looker.ExecutableCommand;
using KPLN_Looker.Services;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Looker
{
    public class Module : IExternalModule
    {
        public static readonly Queue<IExecutableCommand> CommandQueue = new Queue<IExecutableCommand>();
        private readonly DBWorkerService _dBWorkerService;

        public Module()
        {
            _dBWorkerService = new DBWorkerService();
        }

        public Result Close()
        {
            return Result.Succeeded;
        }

        public Result Execute(UIControlledApplication application, string tabName)
        {
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
                if (_dBWorkerService.CurrentDBUser.SubDepartmentId != 8)
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
            string fileName = doc.PathName;
            if (IsMonitoredFile(doc, fileName))
            {
                string centralPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath());
                #region Отлов пользователей с ограничением допуска к работе
                if (_dBWorkerService.CurrentDBUser.IsUserRestricted)
                {
                    BitrixMessageSender.SendErrorMsg_ToBIMChat(
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
                        // Вывожу окно, если джокумент не связан с проектом из БД
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
                    }
                }
                else
                {
                    TaskDialog td = new TaskDialog("ВНИМАНИЕ")
                    {
                        MainIcon = TaskDialogIcon.TaskDialogIconInformation,
                        MainInstruction = "Вы работаете в незарегестрированном проекте. Скинь скрин в BIM-отдел",
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
            if (activeView != null
                && activeView is View3D _
                && (activeView.Title.ToUpper().Contains("BIM360")
                    || activeView.Title.ToUpper().Contains("NAVISWORKS")
                    || activeView.Title.ToUpper().Contains("NWC")
                    || activeView.Title.ToUpper().Contains("NWD"))
                && !_dBWorkerService.CurrentDBUserSubDepartment.Code.ToUpper().Contains("BIM")
                )
            {
                TaskDialog td = new TaskDialog("ВНИМАНИЕ")
                {
                    MainIcon = TaskDialogIcon.TaskDialogIconError,
                    MainInstruction = "Данный вид предназначен только для bim-отдела. Его запрещено открывать или редактировать. Вид будет закрыт",
                    CommonButtons = TaskDialogCommonButtons.Ok,
                };
                td.Show();

                CommandQueue.Enqueue(new ViewCloser(activeView.Id));
            }
            #endregion
        }

        /// <summary>
        /// Событие на изменение в документе
        /// </summary>
        private void OnDocumentChanged(object sender, DocumentChangedEventArgs args)
        {
            var a = 1;
            //try
            //{
            //    #region Анализ триггерных изменений по проекту
            //    Document doc = args.GetDocument();
            //    // Игнорирую не для совместной работы
            //    if (doc.IsWorkshared)
            //    {
            //        // Игнорирую файлы не с диска Y: и файлы концепции
            //        string docPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath());
            //        if (docPath.Contains("stinproject.local\\project\\")
            //            && !(docPath.ToLower().Contains("кон") || docPath.ToLower().Contains("kon")))
            //        {
            //            IsFamilyLoadedFromOtherFile(args);
            //        }
            //    }
            //    #endregion

            //    if (FileActivityService.ActiveDocument != null && !FileActivityService.ActiveDocument.IsFamilyDocument && !FileActivityService.ActiveDocument.PathName.Contains(".rte"))
            //    {
            //        ActivityInfo info = new ActivityInfo(FileActivityService.ActiveDocument, Collections.BuiltInActivity.DocumentChanged);
            //        FileActivityService.ActivityBag.Enqueue(info);
            //    }
            //}
            //catch (Exception) { }
        }

        /// <summary>
        /// Событие на синхронизацию файла
        /// </summary>
        private void OnDocumentSynchronized(object sender, DocumentSynchronizedWithCentralEventArgs args)
        {
            Document doc = args.Document;
            string fileName = doc.PathName;
            if (IsMonitoredFile(doc, fileName))
            {
                string centralPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath());
                #region Отлов пользователей с ограничением допуска к работе
                if (_dBWorkerService.CurrentDBUser.IsUserRestricted)
                {
                    BitrixMessageSender.SendErrorMsg_ToBIMChat(
                        $"Сотрудник: {_dBWorkerService.CurrentDBUser.Surname} {_dBWorkerService.CurrentDBUser.Name} из отдела {_dBWorkerService.CurrentDBUserSubDepartment.Code}\n" +
                        $"Статус допуска: Ограничен в работе с реальными проектами (IsUserRestricted={_dBWorkerService.CurrentDBUser.IsUserRestricted})\n" +
                        $"Действие: Произвел синхронизацию проекта {doc.Title}.");

                    MessageBox.Show(
                        $"Вы произвели синхронизацию проекта с диска Y:\\. Данные переданы в BIM-отдел и ГИ бюро",
                        "Ошибка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
                #endregion

                #region Работа с проектами КПЛН
                DBProject dBProject = _dBWorkerService.Get_DBProjectByRevitDocFile(centralPath);
                if (dBProject != null)
                {
                    DBDocument dBDocument = _dBWorkerService.Get_DBDocumentByRevitDocPathAndDBProject(centralPath, dBProject);
                    if (dBDocument != null)
                    {
                        // Защита закрытого проекта от изменений
                        if (dBDocument.IsClosed)
                        {
                            BitrixMessageSender
                                .SendErrorMsg_ToBIMChat($"Сотрудник: {_dBWorkerService.CurrentDBUser.Surname} {_dBWorkerService.CurrentDBUser.Name} из отдела {_dBWorkerService.CurrentDBUserSubDepartment.Code}\n" +
                                $"Действие: Синхронизиция в проекте {doc.Title}, который ЗАКРЫТ.");

                            TaskDialog taskDialog = new TaskDialog("KPLN: Закрытый проект")
                            {
                                MainIcon = TaskDialogIcon.TaskDialogIconError,
                                MainContent = "Факт синхронизации в закрытом проекте - передан в BIM-отдел",
                                FooterText = "За тобой уже выехали :)",
                                CommonButtons = TaskDialogCommonButtons.Ok,
                            };
                            taskDialog.Show();
                        }
                        _dBWorkerService.Update_DBDocumentLastChangedData(dBDocument);
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
            //// Игнорирую не для совместной работы
            //if (!args.Document.IsWorkshared)
            //    return;

            //Application app = sender as Application;
            //Document prjDoc = args.Document;
            //string familyName = args.FamilyName;
            //string familyPath = args.FamilyPath;

            //// Игнорирую файлы не с диска Y, файлы концепции
            //string docPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(prjDoc.GetWorksharingCentralModelPath());
            //if (!docPath.Contains("stinproject.local\\project\\")
            //    || docPath.ToLower().Contains("конц"))
            //    return;

            //// Отлов семейств которые могут редактировать проектировщики
            //DocumentSet appDocsSet = app.Documents;
            //foreach (Document doc in appDocsSet)
            //{
            //    if (doc.Title.Contains($"{familyName}"))
            //    {
            //        if (doc.IsFamilyDocument)
            //        {
            //            Family family = doc.OwnerFamily;
            //            Category famCat = family.FamilyCategory;
            //            BuiltInCategory bic = (BuiltInCategory)famCat.Id.IntegerValue;

            //            // Отлов семейств марок (могут разрабатывать все)
            //            if (bic.Equals(BuiltInCategory.OST_ProfileFamilies)
            //                || bic.Equals(BuiltInCategory.OST_DetailComponents)
            //                || bic.Equals(BuiltInCategory.OST_GenericAnnotation)
            //                || bic.Equals(BuiltInCategory.OST_DetailComponentsHiddenLines)
            //                || bic.Equals(BuiltInCategory.OST_DetailComponentTags))
            //                return;

            //            // Отлов семейств марок (могут разрабатывать все), за исключением штампов, подписей и жуков
            //            if (famCat.CategoryType.Equals(CategoryType.Annotation)
            //                && !familyName.StartsWith("020_")
            //                && !familyName.StartsWith("022_")
            //                && !familyName.StartsWith("023_")
            //                && !familyName.ToLower().Contains("жук"))
            //                return;

            //            // Отлов семейств лестничных маршей и площадок, которые по форме зависят от проектов (могут разрабатывать все)
            //            if (bic.Equals(BuiltInCategory.OST_GenericModel)
            //                && (familyName.StartsWith("208_") || familyName.StartsWith("209_")))
            //                return;

            //            // Отлов семейств соед. деталей каб. лотков производителей: Ostec, Dkc
            //            if (bic.Equals(BuiltInCategory.OST_CableTrayFitting)
            //                && (familyName.ToLower().Contains("ostec") || familyName.ToLower().Contains("dkc")))
            //                return;
            //        }
            //        else
            //            throw new Exception("Ошибка определения типа файла. Обратись к разработчику!");
            //    }
            //}

        }

        /// <summary>
        /// Проверка файла на наличие мониторинга (проверок)
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private bool IsMonitoredFile(Document doc, string fileName) =>
            doc.IsWorkshared
            && !doc.IsDetached
            && !doc.IsFamilyDocument
            && !fileName.EndsWith("rte")
            && !fileName.ToLower().Contains("\\lib\\")
            && !fileName.ToLower().Contains("концепция");

    }
}
