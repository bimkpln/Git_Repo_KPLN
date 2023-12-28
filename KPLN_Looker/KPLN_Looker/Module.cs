using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using KPLN_Loader.Common;
using KPLN_Looker.Services;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Looker
{
    public class Module : IExternalModule
    {
        private bool _isDocumentClosed = false;
        private readonly UserDbService _userDbService;
        private readonly DBUser _dBUser;
        private readonly DocumentDbService _documentDbService;
        private readonly ProjectDbService _projectDbService;
        private readonly SubDepartmentDbService _subDepartmentDbService;

        public Module()
        {
            _userDbService = (UserDbService)new CreatorUserDbService().CreateService();
            _dBUser = _userDbService.GetCurrentDBUser();

            _documentDbService = (DocumentDbService)new CreatorDocumentDbService().CreateService();
            _projectDbService = (ProjectDbService)new CreatorProjectDbService().CreateService();
            _subDepartmentDbService = (SubDepartmentDbService)new CreatorSubDepartmentDbService().CreateService();
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
                INIFileService iNIFileService = new INIFileService(_dBUser, application.ControlledApplication.VersionNumber);
                if (!iNIFileService.OverwriteINIFile())
                {
                    throw new Exception($"Ошибка при перезаписи ini-файла");
                }

                //Подписка на события
                application.ControlledApplication.DocumentOpened += OnDocumentOpened;
                application.ViewActivated += OnViewActivated;
                application.ControlledApplication.DocumentChanged += OnDocumentChanged;
                application.ControlledApplication.DocumentSynchronizedWithCentral += OnDocumentSynchronized;
                if (_dBUser.SubDepartmentId != 8)
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
        /// Событие на открытие документа
        /// </summary>
        private void OnDocumentOpened(object sender, DocumentOpenedEventArgs args)
        {
            Document doc = args.Document;
            FileInfo fileInfo = new FileInfo(ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath()));
            string fileName = fileInfo.FullName;

            DBDocument dBDocument = null;
            if (doc.IsWorkshared
                && !doc.IsDetached
                && !doc.IsFamilyDocument
                && !fileInfo.Extension.Equals("rte")
                && !fileName.ToLower().Contains(".концепция"))
            {
                DBProject dBProject = _projectDbService.GetDBProjects().Where(p => fileName.Contains(p.MainPath)).FirstOrDefault();
                DBSubDepartment dBSubDepartment = _subDepartmentDbService.GetDBSubDepartment_ByRevitDoc(doc);
                int dBProjectId = dBProject == null ? -1 : dBProject.Id;
                int dBSubDepartmentId = dBSubDepartment == null ? -1 : dBSubDepartment.Id;

                dBDocument = _documentDbService.GetDBDocuments_ByPrjIdAndSubDepId(dBProjectId, dBSubDepartmentId).Where(d => d.FullPath.Equals(fileName)).FirstOrDefault();
                if (dBDocument == null)
                {
                    dBDocument = new DBDocument()
                    {
                        Name = doc.Title,
                        FullPath = fileName,
                        ProjectId = dBProject.Id,
                        SubDepartmentId = dBSubDepartment.Id,
                        IsClosed = false,
                    };
                    Task createNewDoc = Task.Run(() =>
                    {
                        _documentDbService.CreateDBDocument(dBDocument);
                    });
                }
                else
                {
                    if (dBProject != null)
                    {
                        // Технический блок: запись статуса документа IsClosed по статусу проекта
                        Task updateDbDoc = Task.Run(() =>
                        {
                            _documentDbService.UpdateDBDocument_IsClosedByProject(dBProject);
                        });
                    }

                    if ((dBProject != null && dBProject.IsClosed) || dBDocument.IsClosed)
                    {
                        _isDocumentClosed = true;
                        TaskDialog taskDialog = new TaskDialog("KPLN: Закрытый проект")
                        {
                            MainIcon = TaskDialogIcon.TaskDialogIconError,
                            MainContent = "Вы пытаетесь работать в закрытом проекте. О факте синхранизации узнает BIM-отдел. " +
                                "Чтобы получить доступ на внесение изменений в этот проект - обратитесь в BIM-отдел",
                            CommonButtons = TaskDialogCommonButtons.Ok,
                        };
                        taskDialog.Show();
                    }
                    else
                        _isDocumentClosed = false;
                }
            }
        }

        /// <summary>
        /// Событие на активацию вида
        /// </summary>
        private void OnViewActivated(object sender, ViewActivatedEventArgs args)
        {
            //try
            //{
            //    if (args.Document.Title != null && !args.Document.IsFamilyDocument)
            //    {
            //        FileActivityService.ActiveDocument = args.Document;
            //    }
            //}
            //catch (Exception)
            //{
            //    FileActivityService.ActiveDocument = null;
            //}
        }

        /// <summary>
        /// Событие на изменение в документе
        /// </summary>
        private void OnDocumentChanged(object sender, DocumentChangedEventArgs args)
        {
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

            // Защита закрытого проекта от изменений
            if (_isDocumentClosed)
            {
                BitrixMessageService
                    .SendErrorMsg_ToBIMChat($"Сотрудник: {_dBUser.Surname} {_dBUser.Name} из отдела {_subDepartmentDbService.GetDBSubDepartment_ByDBUser(_dBUser).Code}\n" +
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
    }
}
