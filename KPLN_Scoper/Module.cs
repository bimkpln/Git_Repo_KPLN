using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using KPLN_Loader.Common;
using KPLN_Scoper.Common;
using KPLN_Scoper.Services;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using static KPLN_Loader.Output.Output;

namespace KPLN_Scoper
{
    public class Module : IExternalModule
    {
        /// <summary>
        /// Информация о пользователе (кэширование)
        /// </summary>
        private readonly SQLUserInfo _dbUserInfo;

        public Module()
        {
            _dbUserInfo = KPLN_Loader.Preferences.User;
        }
        
        public Result Close()
        {
            try
            {
                FileActivityService.Synchronize(null);
                FileActivityService.Destroy();
            }
            catch (ArgumentNullException ex)
            {
                PrintError(ex);
            }
            
            return Result.Succeeded;
        }
        
        public Result Execute(UIControlledApplication application, string tabName)
        {
            try
            {
                // Перезапись ini-файла
                INIFileService iNIFileService = new INIFileService(_dbUserInfo, application.ControlledApplication.VersionNumber);
                if (!iNIFileService.OverwriteINIFile())
                {
                    throw new Exception($"Ошибка при перезаписи ini-файла");
                }

                // Обновление файлов в БД.
                KPLN_Library_DataBase.DbControll.Update();

                UpdateAllDocumentInfo();

                UpdateAllModuleInfo();

                //Подписка на события
                application.ViewActivated += OnViewActivated;
                application.ControlledApplication.DocumentOpened += OnDocumentOpened;
                application.ControlledApplication.DocumentSynchronizedWithCentral += OnDocumentSynchronized;
                application.ControlledApplication.DocumentChanged += OnDocumentChanged;

                if (!_dbUserInfo.Department.Code.Equals("BIM"))
                {
                    application.ControlledApplication.FamilyLoadingIntoDocument += OnFamilyLoadingIntoDocument;
                } 

                try
                {
                    FileActivityService.Run();
                }
                catch (Exception e)
                {
                    Print($"Ошибка запуске FileActivityService: {e.Message}", KPLN_Loader.Preferences.MessageType.Error);
                }
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                Print($"Ошибка при инициализации: {ex.Message}", KPLN_Loader.Preferences.MessageType.Error);
                return Result.Failed;
            }
        }
        
        private void OnViewActivated(object sender, ViewActivatedEventArgs args)
        {
            try
            {
                if (args.Document.Title != null && !args.Document.IsFamilyDocument )
                {
                    FileActivityService.ActiveDocument = args.Document;
                }
            }
            catch (Exception)
            {
                FileActivityService.ActiveDocument = null;
            }
        }
        
        private void OnDocumentChanged(object sender, DocumentChangedEventArgs args)
        {
            try
            {
                #region Анализ триггерных изменений по проекту
                Document doc = args.GetDocument();
                // Игнорирую не для совместной работы
                if (doc.IsWorkshared)
                {
                    // Игнорирую файлы не с диска Y: и файлы концепции
                    string docPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath());
                    if (docPath.Contains("stinproject.local\\project\\")
                        && !(docPath.ToLower().Contains("кон") || docPath.ToLower().Contains("kon")))
                    {
                        IsFamilyLoadedFromOtherFile(args);
                    }
                }
                #endregion

                if (FileActivityService.ActiveDocument != null && !FileActivityService.ActiveDocument.IsFamilyDocument && !FileActivityService.ActiveDocument.PathName.Contains(".rte"))
                {
                    ActivityInfo info = new ActivityInfo(FileActivityService.ActiveDocument, Collections.BuiltInActivity.DocumentChanged);
                    FileActivityService.ActivityBag.Enqueue(info);
                }
            }
            catch (Exception) { }
        }

        /// <summary>
        /// Проверка семейств на загрузку из другого проекта
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
                        {
                            addedFamilySymbols.Add(familySymbol);
                        }
                    }
                }

                if (addedFamilySymbols.Count() > 0)
                {
                    FilteredElementCollector prjFamilies = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).WhereElementIsElementType();
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
                        TaskDialog.Show("Предупреждение",
                            "Только что были скопированы семейства, которые являются как новыми, так и уже имеющимися в проекте. " +
                            "Запусти плагин KPLN для проверки семейств");
                    else if (isFamilyInclude)
                        TaskDialog.Show("Предупреждение",
                            "Только что были скопированы семейства, которые уже имеющимися в проекте. " +
                            "Запусти плагин KPLN для проверки семейств, чтобы избежать дублирования семейств");
                    else if (isFamilyNew)
                        TaskDialog.Show("Предупреждение",
                            "Только что были скопированы семейства, которые являются новыми. " +
                            "Запусти плагин KPLN для проверки семейств, чтобы избежать наличия семейств из сторонних источников");
                }
            }
        }

        /// <summary>
        /// Контроль процесса загрузки семейств в проекты КПЛН
        /// </summary>
        private void OnFamilyLoadingIntoDocument(object sender, FamilyLoadingIntoDocumentEventArgs args)
        {
            // Игнорирую не для совместной работы
            if (!args.Document.IsWorkshared)
                return;

            Application app = sender as Application;
            Document prjDoc = args.Document;
            string familyName = args.FamilyName;
            string familyPath = args.FamilyPath;
            
            // Игнорирую файлы не с диска Y, файлы концепции
            string docPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(prjDoc.GetWorksharingCentralModelPath());
            if (!docPath.Contains("stinproject.local\\project\\")
                || docPath.ToLower().Contains("конц"))
                return;

            // Отлов семейств которые могут редактировать проектировщики
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

            // Отлов семейств, расположенных не на Х, не из плагинов и не из исключений выше
            if (!familyPath.StartsWith("X:\\BIM") 
                && !familyPath.Contains("KPLN_Loader"))
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
        }

        private void OnDocumentSynchronized(object sender, DocumentSynchronizedWithCentralEventArgs args)
        {
            try
            {
                if (FileActivityService.ActiveDocument != null)
                {
                    ActivityInfo info = new ActivityInfo(FileActivityService.ActiveDocument, Collections.BuiltInActivity.DocumentSynchronized);
                    FileActivityService.ActivityBag.Enqueue(info);
                }
            }
            catch (Exception) { }
            /*
            try
            {
                Autodesk.Revit.DB.Document document = args.Document;
                UpdateRoomDictKeys(document);
            }
            catch (Exception)
            { }
            */
            try
            {
                FileActivityService.Synchronize(null);
            }
            catch (ArgumentNullException ex)
            {
                PrintError(ex);
            }
            catch (Exception)
            { }
        }

        /// <summary>
        /// Событие на открытие документа
        /// </summary>
        private void OnDocumentOpened(object sender, DocumentOpenedEventArgs args)
        {
            try
            {
                if (FileActivityService.ActiveDocument != null 
                    && !args.Document.IsFamilyDocument 
                    && !args.Document.PathName.Contains(".rte"))
                {
                    ActivityInfo info = new ActivityInfo(FileActivityService.ActiveDocument, Collections.BuiltInActivity.ActiveDocument);
                    FileActivityService.ActivityBag.Enqueue(info);

                    // Если отловить ошибку в ActivityInfo - активность по проекту не будет писаться вовсе
                    if (info.ProjectId == -1
                        && info.DocumentId != -1
                        && !info.DocumentTitle.ToLower().Contains("конц"))
                    {
                        Print($"Внимание: Ваш проект не зарегестрирован! Если это временный файл" +
                            " - можете продолжить работу. Если же это файл новго проекта - напишите " +
                            "руководителю BIM-отдела",
                            KPLN_Loader.Preferences.MessageType.Error);
                    }
                }
            }
            catch (Exception ex) 
            { 
                PrintError(ex); 
            }

            try
            {
                Document document = args.Document;

                if (document.IsWorkshared && !document.IsDetached)
                {
                    DocumentPreapre(document);
                }

                foreach (RevitLinkInstance link in new FilteredElementCollector(document).OfClass(typeof(RevitLinkInstance)).WhereElementIsNotElementType().ToElements())
                {
                    Document linkDocument = link.GetLinkDocument();
                    if (linkDocument == null) { continue; }

                    if (linkDocument.IsWorkshared)
                    {
                        DocumentPreapre(linkDocument);
                    }
                }
            }
            catch (Exception) { }
        }


        /// <summary>
        /// Обновление данных по документам в БД администратором БД (по имени пользователя)
        /// </summary>
        private void UpdateAllDocumentInfo()
        {
            if (_dbUserInfo.SystemName != "tkutsko")
            { return; }
            
            List<SQLProject> projects = GetProjects();
            List<SQLDepartment> departments = GetDepartments();
            List<SQLDocument> documents = GetDocuments();
            foreach (SQLDocument doc in documents)
            {
                // Проверка на присутсвие файла на сервере. Игнорируется revit-server, пустые имена
                bool isExist = true && (doc.Path.Contains("RSN") || (!doc.Path.Equals(String.Empty) && new FileInfo(doc.Path).Exists));
                if (isExist)
                {
                    SQLiteConnection sql = new SQLiteConnection(KPLN_Library_DataBase.DbControll.MainDBConnection);
                    // Анализ кодов документов из БД
                    try
                    {
                        // Игнорирую файлы ревизора с диска С, файлы концепции, отсоединенные файлы
                        if (
                            doc.Path.Contains($"C:\\") 
                            || doc.Name.ToLower().Contains("отсоединено") 
                            || doc.Path.ToLower().Contains("концепц") 
                            || doc.Name.ToLower().Contains("_кон_") 
                            || doc.Name.ToLower().Contains("_kon_"))
                        {
                            continue;
                        }

                        // Поиск проекта и отдела по id
                        SQLProject pickedProject = GetProjectById(projects, doc.Project);
                        SQLDepartment pickedDepartment = GetDepartmentById(departments, doc.Department);

                        // Поиск отдела и проекта по ключам (часть пути файла и имя файла)
                        if (doc.Department == -1 || doc.Project == -1)
                        {
                            // Поиск проекта и отдела по пути или имени
                            pickedProject = GetProjectByPath(projects, doc.Path);
                            pickedDepartment = GetDepartmentByDocName(departments, doc.Name);

                            if (pickedProject == null || pickedDepartment == null)
                            {
                                Print($"Проблемы с привязкой к проекту и отделу в указанном документе (Documents): " +
                                    $"Id: {doc.Id}, Name: {doc.Name}, Department: {doc.Department}, Project: {doc.Project}, Code: {doc.Code}",
                                    KPLN_Loader.Preferences.MessageType.Error);
                                continue;
                            }

                            // Обновление данных по проекту и отделу исходя из ключей
                            sql.Open();
                            SQLiteCommand cmd = new SQLiteCommand($"UPDATE Documents SET Department = '{pickedDepartment.Id}' WHERE Id = {doc.Id}; " +
                                $"UPDATE Documents SET Project = '{pickedProject.Id}' WHERE Id = {doc.Id}; ", sql);
                            cmd.ExecuteNonQuery();
                            sql.Close();
                        }

                        // Назначение шифра проекта по коду и отделу
                        if (doc.Code == "NONE")
                        {
                            sql.Open();
                            string code = $"{pickedProject.Code}_{pickedDepartment.Code}";
                            SQLiteCommand cmd = new SQLiteCommand($"UPDATE Documents SET Code = '{code}' WHERE Id = {doc.Id};", sql);
                            cmd.ExecuteNonQuery();
                            sql.Close();
                        }
                    }
                    finally { sql.Close(); }
                }
                else
                {
                    SQLiteConnection db = new SQLiteConnection(KPLN_Library_DataBase.DbControll.MainDBConnection);
                    try
                    {
                        db.Open();
                        SQLiteCommand cmd_insert = new SQLiteCommand(string.Format("DELETE FROM Documents WHERE Path='{0}'", doc.Path), db);
                        cmd_insert.ExecuteNonQuery();
                        db.Close();
                    }
                    finally { db.Close(); }
                }
            }

            Print($"Автоопределение кодов документов БД выполнено!", KPLN_Loader.Preferences.MessageType.Warning);
        }
        /*
        private void UpdateRoomDictKeys(Document doc)
        {
            if (new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements().Count == 0)
            {
                return;
            }
            HashSet<string> room_names = new HashSet<string>();
            HashSet<string> room_departments = new HashSet<string>();
            foreach (Room room in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements())
            {
                try
                {
                    room_names.Add(room.get_Parameter( BuiltInParameter.ROOM_NAME).AsString().ToLower());
                }
                catch (Exception)
                { }
                try
                {
                    string dep = room.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT).AsString().ToLower();
                    if (dep != string.Empty)
                    {
                        room_departments.Add(dep);
                    }
                }
                catch (Exception)
                { }
            }
            Thread t = new Thread(() =>
            {
                SQLiteConnection sql = new SQLiteConnection();
                try
                {
                    sql.ConnectionString = string.Format(@"Data Source=Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_Dictionary.db;Version=3;");
                    sql.Open();
                    HashSet<string> stored_names = new HashSet<string>();
                    HashSet<string> stored_departments = new HashSet<string>();
                    using (SQLiteCommand cmd = new SQLiteCommand("SELECT KEY FROM DICT_ROOMS_NAMES", sql))
                    {
                        using (SQLiteDataReader rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                stored_names.Add(rdr.GetString(0));
                            }
                        }
                    }
                    using (SQLiteCommand cmd = new SQLiteCommand("SELECT KEY FROM DICT_ROOMS_DEPARTMENTS", sql))
                    {
                        using (SQLiteDataReader rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                stored_departments.Add(rdr.GetString(0));
                            }
                        }
                    }
                    foreach (string value in room_names)
                    {
                        if (!stored_names.Contains(value))
                        {
                            SQLiteCommand cmd_insert = new SQLiteCommand("INSERT INTO DICT_ROOMS_NAMES ([KEY], [TYPE]) VALUES(@KEY, @TYPE)", sql);
                            cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "KEY", Value = value });
                            cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "TYPE", Value = -1 });
                            cmd_insert.ExecuteNonQuery();
                        }
                    }
                    foreach (string value in room_departments)
                    {
                        if (!stored_departments.Contains(value))
                        {
                            SQLiteCommand cmd_insert = new SQLiteCommand("INSERT INTO DICT_ROOMS_DEPARTMENTS ([KEY], [TYPE]) VALUES(@KEY, @TYPE)", sql);
                            cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "KEY", Value = value });
                            cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "TYPE", Value = -1 });
                            cmd_insert.ExecuteNonQuery();
                        }
                    }
                    sql.Close();
                }
                catch (Exception) { }
                {
                    try
                    {
                        sql.Close();
                    }
                    catch (Exception) { }
                }
            });
            t.Start();
        }
        */

        /// <summary>
        /// Обновление данных по модулям в БД администратором БД (по имени пользователя)
        /// </summary>
        private void UpdateAllModuleInfo()
        {
            if (_dbUserInfo.SystemName != "tkutsko")
            { return; }

            List<SQLModuleInfo> modules = GetModules();
            foreach (SQLModuleInfo module in modules) 
            {
                bool isEdit = false;
                string moduleName = module.Name;
                string modulePath = module.Path;

                DirectoryInfo directoryInfo = new DirectoryInfo(modulePath);
                foreach (FileInfo file in directoryInfo.GetFiles())
                {
                    if (file.Name.Contains(moduleName))
                    {
                        FileVersionInfo moduleFileVersionInfo = FileVersionInfo.GetVersionInfo($"{modulePath}\\{file.Name}");

                        string dllVersion = moduleFileVersionInfo.FileVersion;
                        if (dllVersion == null)
                            continue;

                        SQLiteConnection sql = new SQLiteConnection(
                            KPLN_Library_DataBase.DbControll.MainDBConnection);
                        
                        try
                        {
                            sql.Open();
                            SQLiteCommand cmd = new SQLiteCommand(
                                string.Format("UPDATE Modules SET Version = '{0}' WHERE Id = {1}",
                                dllVersion, 
                                module.Id), sql);
                            
                            cmd.ExecuteNonQuery();
                            isEdit = true;
                        }
                        catch (Exception ex) 
                        {
                            PrintError(ex);
                        }
                        finally
                        {
                            sql.Close();
                        }
                    }
                }

                if (!isEdit && !moduleName.Contains("Test"))
                {
                    Print($"Проблемы с обновлением версии палгина " +
                        $"Id: {module.Id}, Name: {moduleName}", KPLN_Loader.Preferences.MessageType.Error);
                }
            }

            Print($"Обновление версий модулей в БД выполнено!", KPLN_Loader.Preferences.MessageType.Warning);
        }

        /// <summary>
        /// Поиск проекта по id
        /// </summary>
        /// <param name="list">Коллекция проектов из БД</param>
        /// <param name="id">Id проекта</param>
        /// <returns></returns>
        private SQLProject GetProjectById(List<SQLProject> list, int id)
        {
            foreach (SQLProject project in list)
            {
                if (project.Id == id)
                {
                    return project;
                }
            }
            return null;
        }

        /// <summary>
        /// Поиск проекта по пути
        /// </summary>
        /// <param name="list">Коллекция проектов из БД</param>
        /// <param name="path">Путь проекта</param>
        /// <returns></returns>
        private SQLProject GetProjectByPath(List<SQLProject> list, string path)
        {
            foreach (SQLProject project in list)
            {
                if (path.Contains(project.Keys))
                {
                    return project;
                }
            }
            return null;
        }

        /// <summary>
        /// Поиск отдела по id
        /// </summary>
        /// <param name="list">Коллекция отделов из БД</param>
        /// <param name="id">Id отдела</param>
        /// <returns></returns>
        private SQLDepartment GetDepartmentById(List<SQLDepartment> list, int id)
        {
            foreach (SQLDepartment department in list)
            {
                if (department.Id == id)
                {
                    return department;
                }
            }
            return null;
        }

        /// <summary>
        /// Поиск отдела по части имени имени
        /// </summary>
        /// <param name="list">Коллекция отделов из БД</param>
        /// <param name="namePart">Часть имени отдела</param>
        /// <returns></returns>
        private SQLDepartment GetDepartmentByNameContain(List<SQLDepartment> list, string namePart)
        {
            foreach (SQLDepartment department in list)
            {
                if (department.Name.ToLower().Contains(namePart))
                {
                    return department;
                }
            }
            return null;
        }

        /// <summary>
        /// Поиск отдела по имени файла
        /// </summary>
        /// <param name="list">Коллекция отделов из БД</param>
        /// <param name="docName">Имя файла</param>
        /// <returns></returns>
        private SQLDepartment GetDepartmentByDocName(List<SQLDepartment> list, string docName)
        {
            // Тонкая отработка разбивочных файлов
            if (docName.ToLower().Contains("разбив")
                || docName.ToLower().Contains("разбфайл")
                || docName.ToLower().Contains("разб.файл"))
            {
                return GetDepartmentByNameContain(list, "разбивочный");
            }
            // Тонкая отработка файлов КР (КЖ)
            else if (docName.ToLower().Contains("кж")
                || docName.ToLower().Contains("kg"))
            {
                return GetDepartmentByNameContain(list, "конструктив");
            }
            // Тонкая отработка файлов ВК (АУПТ)
            else if (docName.ToLower().Contains("пт")
                || docName.ToLower().Contains("аупт")
                || docName.ToLower().Contains("pt")
                || docName.ToLower().Contains("aupt"))
            {
                return GetDepartmentByNameContain(list, "система пожаротушения");
            }
            // Тонкая отработка файлов ОВ (ИТП)
            else if (docName.ToLower().Contains("итп")
                || docName.ToLower().Contains("itp"))
            {
                return GetDepartmentByNameContain(list, "тепловой пункт");
            }
            // Тонкая отработка файлов СС (автоматизация)
            else if (docName.ToLower().Contains("ав")
                || docName.ToLower().Contains("av"))
            {
                return GetDepartmentByNameContain(list, "автоматизация");
            }
            // Обрабатываю остальные файлы
            else
            {
                foreach (SQLDepartment department in list)
                {
                    if (docName.Contains($"_{department.Code}") 
                        || docName.Contains($"_{department.CodeUs}")
                        || docName.Contains($"{department.Code}_") 
                        || docName.Contains($"{department.CodeUs}_")
                        || docName.Contains($"-{department.Code}")
                        || docName.Contains($"-{department.CodeUs}")
                        || docName.Contains($"{department.Code}-")
                        || docName.Contains($"{department.CodeUs}-"))
                    {
                        return department;
                    }
                }
            }
            return null;
        }

        private bool AddDocument(SQLDocument document)
        {
            SQLiteConnection sql = new SQLiteConnection(KPLN_Library_DataBase.DbControll.MainDBConnection);
            try
            {
                sql.Open();
                using (SQLiteCommand cmd = sql.CreateCommand())
                {
                    cmd.CommandText = "INSERT INTO Documents ([Path], [Name], [Department], [Project], [Code]) VALUES (@Path, @Name, @Department, @Project, @Code)";
                    cmd.Parameters.Add(new SQLiteParameter() { ParameterName = "@Path", Value = document.Path });
                    cmd.Parameters.Add(new SQLiteParameter() { ParameterName = "@Name", Value = document.Name });
                    cmd.Parameters.Add(new SQLiteParameter() { ParameterName = "@Department", Value = document.Department });
                    cmd.Parameters.Add(new SQLiteParameter() { ParameterName = "@Project", Value = document.Project });
                    cmd.Parameters.Add(new SQLiteParameter() { ParameterName = "@Code", Value = document.Code });
                    cmd.ExecuteNonQuery();
                }
                sql.Close();
                return true;
            }
            catch (Exception) { }
            {
                try
                {
                    sql.Close();
                }
                catch (Exception) { }

            }
            return false;
        }
        
        private List<SQLDocument> GetDocuments()
        {
            List<SQLDocument> documents = new List<SQLDocument>();
            SQLiteConnection sql = new SQLiteConnection(KPLN_Library_DataBase.DbControll.MainDBConnection);
            try
            {
                sql.Open();
                using (SQLiteCommand cmd = new SQLiteCommand("SELECT * FROM Documents", sql))
                {
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            if (rdr.GetInt32(0) == -1) { continue; }
                            documents.Add(new SQLDocument(
                                rdr.GetInt32(0), 
                                rdr.GetString(1), 
                                rdr.GetString(2), 
                                rdr.GetInt32(3), 
                                rdr.GetInt32(4), 
                                rdr.GetString(5)));
                        }
                    }
                }
                sql.Close();
            }
            catch (Exception) { }
            {
                try
                {
                    sql.Close();
                }
                catch (Exception) { }
            }
            return documents;
        }

        /// <summary>
        /// Подготовка, проверка и инициализация записи докумнета в БД
        /// </summary>
        private void DocumentPreapre(Document document)
        {
            string path = ModelPathUtils.ConvertModelPathToUserVisiblePath(document.GetWorksharingCentralModelPath());
            if (path == null)
                throw new Exception($"Проблемы с определением пути у файла {document.PathName}. Скинь скрин ошибки в BIM-отдел");

            string name = path.Split(new string[] { @"\" }, StringSplitOptions.None).Last();
            List<string> parts = new List<string>();
            SQLProject pickedProject = null;
            SQLDepartment pickedDepartment = null;

            List<SQLProject> dbProjects = GetProjects();
            foreach (SQLProject prj in dbProjects)
            {
                if (prj.Keys.Contains('*'))
                {
                    foreach (string part in prj.Keys.Split('*'))
                    {
                        if (path != string.Empty && path.Contains(part))
                        {
                            pickedProject = prj;
                        }
                    }
                }
                else
                {
                    if (path.Contains(prj.Keys))
                    {
                        pickedProject = prj;
                    }
                }

            }

            List<SQLDepartment> dbDepartments = GetDepartments();
            // Тонкая отработка файлов БИМ-отдела
            if (pickedProject != null && pickedProject.Id == 1)
            {
                pickedDepartment = dbDepartments.Where(d => d.Id == 10).FirstOrDefault();
            }
            // Отработка остальных файлов
            foreach (SQLDepartment dep in dbDepartments)
            {
                if (name.Contains($"_{dep.Code}") 
                    || name.Contains($"_{dep.CodeUs}")
                    || name.Contains($"{dep.Code}_")
                    || name.Contains($"{dep.CodeUs}_"))
                {
                    pickedDepartment = dep;
                }
            }
            
            int department = -1;
            int project = -1;
            if (pickedProject != null)
            {
                project = pickedProject.Id;
            }
            if (pickedDepartment != null)
            {
                department = pickedDepartment.Id;
            }
            
            string code = "NONE";
            if (pickedProject != null && pickedDepartment != null)
            {
                code = string.Format("{0}_{1}", pickedProject.Code, pickedDepartment.Code);
            }
            
            SQLDocument scopeDocument = new SQLDocument(
                -1, 
                path, 
                name, 
                department, 
                project, 
                code);
            
            bool documentExist = false;
            foreach (SQLDocument doc in GetDocuments())
            {
                if (doc.Path == scopeDocument.Path)
                {
                    documentExist = true;
                }
            }
            if (!documentExist)
            {
                AddDocument(scopeDocument);
            }
        }
        
        private List<SQLProject> GetProjects()
        {
            List<SQLProject> projects = new List<SQLProject>();
            SQLiteConnection sql = new SQLiteConnection(KPLN_Library_DataBase.DbControll.MainDBConnection);
            try
            {
                sql.Open();
                using (SQLiteCommand cmd = new SQLiteCommand("SELECT * FROM Projects", sql))
                {
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            projects.Add(new SQLProject(rdr.GetInt32(0), rdr.GetString(1), rdr.GetString(3), rdr.GetString(4)));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PrintError(ex);
            }
            finally
            {
                sql.Close();
            }
            return projects;
        }
        
        private List<SQLDepartment> GetDepartments()
        {
            List<SQLDepartment> departments = new List<SQLDepartment>();
            SQLiteConnection sql = new SQLiteConnection(KPLN_Library_DataBase.DbControll.MainDBConnection);
            try
            {
                sql.Open();
                using (SQLiteCommand cmd = new SQLiteCommand("SELECT * FROM SubDepartments", sql))
                {
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            departments.Add(new SQLDepartment(rdr.GetInt32(0), rdr.GetString(1), rdr.GetString(2), rdr.GetString(3)));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PrintError(ex);
            }
            finally
            {
                sql.Close();
            }
            return departments;
        }

        /// <summary>
        /// Взять все модули из БД
        /// </summary>
        /// <returns>Коллекция SQLModuleInfo</returns>
        private List<SQLModuleInfo> GetModules()
        {
            List<SQLModuleInfo> modules = new List<SQLModuleInfo>();
            
            SQLiteConnection sql = new SQLiteConnection(KPLN_Library_DataBase.DbControll.MainDBConnection);
            try
            {
                sql.Open();
                using (SQLiteCommand cmd = new SQLiteCommand("SELECT * FROM Modules", sql))
                {
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            modules.Add(new SQLModuleInfo(rdr.GetInt32(0), rdr.GetInt32(1), rdr.GetString(2), rdr.GetString(3), rdr.GetString(4)));
                        }
                    }
                }
            }
            catch (Exception ex) 
            {
                PrintError(ex);
            }
            finally
            {
                sql.Close();
            }

            return modules;
        }
    }
}
