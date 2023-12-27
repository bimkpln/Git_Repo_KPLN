using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using KPLN_Library_SQLiteWorker.FactoryParts.Common;
using KPLN_Loader.Common;
using KPLN_Looker.Services;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Looker
{
    public class Module : IExternalModule
    {
        private readonly string _assemblyPath = Assembly.GetExecutingAssembly().Location;
        private readonly DBUser _dbUser;

        public Module()
        {
            UserDbService userDbService = (UserDbService)new CreatorUserDbService().CreateService();
            _dbUser = userDbService.GetCurrentDBUser();
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
                INIFileService iNIFileService = new INIFileService(_dbUser, application.ControlledApplication.VersionNumber);
                if (!iNIFileService.OverwriteINIFile())
                {
                    throw new Exception($"Ошибка при перезаписи ini-файла");
                }

                //Подписка на события
                application.ControlledApplication.DocumentOpened += OnDocumentOpened;
                application.ViewActivated += OnViewActivated;
                application.ControlledApplication.DocumentChanged += OnDocumentChanged;
                application.ControlledApplication.DocumentSynchronizedWithCentral += OnDocumentSynchronized;
                if (_dbUser.SubDepartmentId != 8)
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
            //try
            //{
            //    if (FileActivityService.ActiveDocument != null
            //        && !args.Document.IsFamilyDocument
            //        && !args.Document.PathName.Contains(".rte"))
            //    {
            //        ActivityInfo info = new ActivityInfo(FileActivityService.ActiveDocument, Collections.BuiltInActivity.ActiveDocument);
            //        FileActivityService.ActivityBag.Enqueue(info);

            //        // Если отловить ошибку в ActivityInfo - активность по проекту не будет писаться вовсе
            //        if (info.ProjectId == -1
            //            && info.DocumentId != -1
            //            && !info.DocumentTitle.ToLower().Contains("конц"))
            //        {
            //            Print($"Внимание: Ваш проект не зарегестрирован! Если это временный файл" +
            //                " - можете продолжить работу. Если же это файл новго проекта - напишите " +
            //                "руководителю BIM-отдела",
            //                KPLN_Loader.Preferences.MessageType.Error);
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    PrintError(ex);
            //}

            //try
            //{
            //    Document document = args.Document;

            //    if (document.IsWorkshared && !document.IsDetached)
            //    {
            //        DocumentPreapre(document);
            //    }

            //    foreach (RevitLinkInstance link in new FilteredElementCollector(document).OfClass(typeof(RevitLinkInstance)).WhereElementIsNotElementType().ToElements())
            //    {
            //        Document linkDocument = link.GetLinkDocument();
            //        if (linkDocument == null) { continue; }

            //        if (linkDocument.IsWorkshared)
            //        {
            //            DocumentPreapre(linkDocument);
            //        }
            //    }
            //}
            //catch (Exception) { }
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
            //try
            //{
            //    if (FileActivityService.ActiveDocument != null)
            //    {
            //        ActivityInfo info = new ActivityInfo(FileActivityService.ActiveDocument, Collections.BuiltInActivity.DocumentSynchronized);
            //        FileActivityService.ActivityBag.Enqueue(info);
            //    }
            //}
            //catch (Exception) { }
            ///*
            //try
            //{
            //    Autodesk.Revit.DB.Document document = args.Document;
            //    UpdateRoomDictKeys(document);
            //}
            //catch (Exception)
            //{ }
            //*/
            //try
            //{
            //    FileActivityService.Synchronize(null);
            //}
            //catch (ArgumentNullException ex)
            //{
            //    PrintError(ex);
            //}
            //catch (Exception)
            //{ }
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
