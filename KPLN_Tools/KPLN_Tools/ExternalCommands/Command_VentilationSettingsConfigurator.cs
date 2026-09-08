using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Tools.Forms;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Command_VentilationSettingsConfigurator : IExternalCommand
    {
        internal const string PluginName = "Конфигуратор вент. установок";
        internal const string UnknownProjectName = "Неизвестно";
        internal const string ProjectDatabasePath =
            @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_Loader_MainDB.db";
        internal const string BaseFamilyTypeName =
            "имяСистемы(при необходимости)_имяУстановкиПоПодборке";
        // Точное имя настроенного типа из исходного RFA. KT в имени — латинские буквы.
        internal const string ConfiguredFamilyTypeName =
            "П6_KT 2,2-250911903-02.02-K-O-P-A";

        // Обычный путь Windows: обратная косая черта перед '_' в сообщении
        // может быть экранированием Markdown, а не разделителем папок.
        internal const string SourceFamilyPath =
            @"X:\BIM\3_Семейства\4_ОВиК\4_ОВ2_Вентиляция\550-596_Инженерное оборудование\550_Универсальная установка_Одноуровневая_(Об).rfa";
        private const string LiteralSourceFamilyPath =
            @"X:\BIM\3\_Семейства\4\_ОВиК\4\_ОВ2\_Вентиляция\550-596\_Инженерное оборудование\550\_Универсальная установка\_Одноуровневая\_(Об).rfa";

        private static VentilationSettingsConfiguratorMain _window;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                if (_window == null || !_window.IsLoaded)
                {
                    var uiapp = commandData.Application;
                    var window = new VentilationSettingsConfiguratorMain(uiapp, uiapp.ActiveUIDocument);
                    _window = window;
                    window.Closed += (s, e) => { if (ReferenceEquals(_window, window)) _window = null; };
                    window.Show();
                }
                else
                {
                    if (_window.WindowState == System.Windows.WindowState.Minimized)
                        _window.WindowState = System.Windows.WindowState.Normal;
                    _window.Activate();
                }
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                if (_window != null && !_window.IsLoaded)
                {
                    _window.ReleaseExternalEvent();
                    _window = null;
                }
                message = ex.Message;
                return Result.Failed;
            }
        }

        internal static string ResolveSourcePath()
        {
            // Сначала проверяем буквальный вариант, затем вариант без Markdown-экранирования.
            return File.Exists(LiteralSourceFamilyPath) ? LiteralSourceFamilyPath : SourceFamilyPath;
        }

        internal enum RequestKind { OpenInRevit, AddToProject }

        internal sealed class OpenDocumentItem
        {
            internal Document Document { get; set; }
            public string DisplayName { get; set; }
            public string LoadTargetName { get; set; }
        }

        internal sealed class ProjectItem
        {
            internal string Name { get; set; }
            internal string Stage { get; set; }
            internal string Code { get; set; }
            internal string[] Paths { get; set; }
            internal bool IsUnknown { get; set; }
            public string DisplayName
            {
                get { return IsUnknown ? UnknownProjectName : Name + " [" + Stage + "]"; }
            }
            public string Description
            {
                get
                {
                    return DisplayName + (Paths == null || Paths.Length == 0
                    ? string.Empty : "\n" + string.Join("\n", Paths));
                }
            }
        }

        internal sealed class ProjectMatch
        {
            internal ProjectItem Project { get; set; }
            internal string RootPath { get; set; }
            internal bool IsAmbiguous { get; set; }
        }

        internal static string CleanDatabaseValue(string value)
        {
            string result = (value ?? string.Empty).Trim();
            return string.Equals(result, "Empty", StringComparison.OrdinalIgnoreCase) ? string.Empty : result;
        }

        // Чтение через SQLite из состава Windows 10/11: дополнительные DLL плагину не нужны.
        // https://www.sqlite.org/c3ref/open.html — SQLITE_OPEN_READONLY не создаёт базу.
        internal static class ProjectDatabase
        {
            private const string Query = "SELECT [Name], [Stage], [Code], [MainPath], " +
                "[RevitServerPath], [RevitServerPath2], [RevitServerPath3], [RevitServerPath4] FROM [Projects];";

            internal static IList<ProjectItem> Read(string path)
            {
                if (!File.Exists(path))
                    throw new FileNotFoundException("База проектов недоступна. Проверьте диск Z: и права на чтение.\n" + path);

                IntPtr database = IntPtr.Zero;
                IntPtr statement = IntPtr.Zero;
                try
                {
                    int result = sqlite3_open_v2(Encoding.UTF8.GetBytes(path + "\0"), out database, 1, IntPtr.Zero);
                    Check(result, database);
                    Check(sqlite3_busy_timeout(database, 1500), database);
                    byte[] sql = Encoding.UTF8.GetBytes(Query + "\0");
                    IntPtr tail;
                    Check(sqlite3_prepare_v2(database, sql, sql.Length, out statement, out tail), database);
                    var projects = new List<ProjectItem>();
                    while ((result = sqlite3_step(statement)) == 100) // SQLITE_ROW
                    {
                        var paths = new List<string>();
                        for (int column = 3; column < 8; column++)
                        {
                            string value = ReadText(statement, column);
                            if (value.Length != 0) paths.Add(value);
                        }
                        string name = ReadText(statement, 0);
                        string stage = ReadText(statement, 1);
                        projects.Add(new ProjectItem
                        {
                            Name = name.Length == 0 ? UnknownProjectName : name,
                            Stage = stage.Length == 0 ? UnknownProjectName : stage,
                            Code = ReadText(statement, 2),
                            Paths = paths.Distinct(StringComparer.OrdinalIgnoreCase).ToArray()
                        });
                    }
                    if (result != 101) Check(result, database); // SQLITE_DONE
                    return projects.OrderBy(p => p.DisplayName, StringComparer.CurrentCultureIgnoreCase).ToList();
                }
                catch (DllNotFoundException ex)
                {
                    throw new InvalidOperationException("Компонент SQLite Windows (winsqlite3.dll) недоступен.", ex);
                }
                finally
                {
                    if (statement != IntPtr.Zero) sqlite3_finalize(statement);
                    if (database != IntPtr.Zero) sqlite3_close(database);
                }
            }

            private static string ReadText(IntPtr statement, int column)
            {
                IntPtr pointer = sqlite3_column_text(statement, column);
                int length = sqlite3_column_bytes(statement, column);
                if (pointer == IntPtr.Zero || length == 0) return string.Empty;
                var bytes = new byte[length];
                Marshal.Copy(pointer, bytes, 0, length);
                return CleanDatabaseValue(Encoding.UTF8.GetString(bytes));
            }

            private static void Check(int result, IntPtr database)
            {
                if (result == 0) return;
                IntPtr pointer = database == IntPtr.Zero ? IntPtr.Zero : sqlite3_errmsg(database);
                var bytes = new List<byte>();
                if (pointer != IntPtr.Zero)
                    for (int i = 0; Marshal.ReadByte(pointer, i) != 0; i++)
                        bytes.Add(Marshal.ReadByte(pointer, i));
                throw new InvalidOperationException("Не удалось прочитать таблицу Projects. SQLite " + result + ": " +
                    Encoding.UTF8.GetString(bytes.ToArray()));
            }

            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
            private static extern int sqlite3_open_v2(byte[] filename, out IntPtr database, int flags, IntPtr vfs);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
            private static extern int sqlite3_busy_timeout(IntPtr database, int milliseconds);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
            private static extern int sqlite3_prepare_v2(IntPtr database, byte[] sql, int length, out IntPtr statement, out IntPtr tail);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
            private static extern int sqlite3_step(IntPtr statement);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
            private static extern IntPtr sqlite3_column_text(IntPtr statement, int column);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
            private static extern int sqlite3_column_bytes(IntPtr statement, int column);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
            private static extern IntPtr sqlite3_errmsg(IntPtr database);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
            private static extern int sqlite3_finalize(IntPtr statement);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
            private static extern int sqlite3_close(IntPtr database);
        }

        // Сопоставление только полных путей и границ папок: «Проект» не совпадает с «Проект2».
        internal static class ProjectPathMatcher
        {
            internal static string Normalize(string value)
            {
                string path = CleanDatabaseValue(value).Replace('\\', '/');
                string prefix;
                int protectedSegments;
                if (path.StartsWith("RSN://", StringComparison.OrdinalIgnoreCase))
                {
                    prefix = "RSN://";
                    path = path.Substring(6);
                    protectedSegments = 1; // имя Revit Server
                }
                else if (path.StartsWith("//", StringComparison.Ordinal))
                {
                    prefix = "//";
                    path = path.Substring(2);
                    protectedSegments = 2; // сервер и общая папка UNC
                }
                else if (path.Length >= 3 && char.IsLetter(path[0]) && path[1] == ':' && path[2] == '/')
                {
                    prefix = path.Substring(0, 3);
                    path = path.Substring(3);
                    protectedSegments = 0;
                }
                else return string.Empty;

                var segments = new List<string>();
                foreach (string segment in path.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    if (segment == ".") continue;
                    if (segment == "..")
                    {
                        if (segments.Count <= protectedSegments) return string.Empty;
                        segments.RemoveAt(segments.Count - 1);
                    }
                    else segments.Add(segment);
                }
                if (segments.Count < protectedSegments) return string.Empty;
                return prefix + string.Join("/", segments);
            }

            internal static ProjectMatch Find(IList<ProjectItem> projects, string modelPath)
            {
                string path = Normalize(modelPath);
                var match = new ProjectMatch();
                if (path.Length == 0) return match;
                int bestLength = -1;
                foreach (var project in projects)
                {
                    foreach (string root in project.Paths ?? new string[0])
                    {
                        string normalizedRoot = Normalize(root);
                        if (normalizedRoot.Length == 0) continue;
                        string folderPrefix = normalizedRoot.TrimEnd('/') + "/";
                        if (!string.Equals(path, normalizedRoot, StringComparison.OrdinalIgnoreCase) &&
                            !path.StartsWith(folderPrefix, StringComparison.OrdinalIgnoreCase)) continue;
                        if (normalizedRoot.Length > bestLength)
                        {
                            bestLength = normalizedRoot.Length;
                            match.Project = project;
                            match.RootPath = root;
                            match.IsAmbiguous = false;
                        }
                        else if (normalizedRoot.Length == bestLength && !ReferenceEquals(match.Project, project))
                            match.IsAmbiguous = true;
                    }
                }
                if (match.IsAmbiguous)
                {
                    match.Project = null;
                    match.RootPath = null;
                }
                return match;
            }
        }

        internal sealed class FamilyRequest
        {
            internal RequestKind Kind { get; set; }
            internal string TypeName { get; set; }
        }

        // Изменения модели проходят через Execute; слежение за активной моделью — через Idling.
        // Классы вложены здесь, чтобы не расширять файловую структуру плагина.
        internal sealed class FamilyRequestHandler : IExternalEventHandler
        {
            private readonly VentilationSettingsConfiguratorMain _owner;
            private UIApplication _application;
            private bool _stopTrackingRequested;
            private IList<ProjectItem> _projects = new List<ProjectItem>();
            private string _lastModelKey;
            internal FamilyRequest PendingRequest { get; set; }

            internal FamilyRequestHandler(VentilationSettingsConfiguratorMain owner)
            {
                _owner = owner;
            }

            public string GetName() { return PluginName; }

            internal void StartTracking(UIApplication app)
            {
                _application = app;
                app.Idling += OnIdling;
                RefreshProjects(app);
            }

            // Можно вызывать из Window.Closed: здесь нет обращений к Revit API.
            internal void RequestStopTracking()
            {
                _stopTrackingRequested = true;
                PendingRequest = null;
            }

            private void OnIdling(object sender, IdlingEventArgs args)
            {
                // Отписка от событий Revit тоже требует API-контекста.
                // Выполняем её здесь, прежде чем читать состояние уже закрытого окна.
                if (_stopTrackingRequested)
                {
                    if (_application != null)
                    {
                        _application.Idling -= OnIdling;
                        _application = null;
                    }
                    return;
                }
                if (_owner.IsBusy || _application == null) return;
                TryUpdateActiveProject(_application, false);
            }

            private void TryUpdateActiveProject(UIApplication app, bool force)
            {
                string filePath = string.Empty;
                try
                {
                    Document document = app.ActiveUIDocument == null ? null : app.ActiveUIDocument.Document;
                    // Читаем фактический путь отдельно от сопоставления проекта, в том числе для RFA.
                    if (document != null && document.IsValidObject)
                        filePath = document.PathName ?? string.Empty;
                    UpdateActiveProject(document, filePath, force);
                }
                catch (Exception ex)
                {
                    // Ошибка чтения пути не должна оставлять выделенным предыдущий проект.
                    string key = "Ошибка определения проекта: " + ex.Message;
                    if (string.Equals(_lastModelKey, key, StringComparison.Ordinal)) return;
                    _lastModelKey = key;
                    _owner.SetProjects(_projects, new ProjectMatch(),
                        string.IsNullOrWhiteSpace(filePath) ? "Путь недоступен" : filePath);
                    _owner.SetStatus(key, true);
                }
            }

            private void RefreshProjects(UIApplication app)
            {
                string databaseError = null;
                try { _projects = ProjectDatabase.Read(ProjectDatabasePath); }
                catch (Exception ex)
                {
                    _projects = new List<ProjectItem>();
                    databaseError = ex.Message;
                }
                TryUpdateActiveProject(app, true);
                _owner.SetStatus(databaseError ?? "Введите имя типа и выберите действие.", databaseError != null);
            }

            private static string[] GetModelPaths(Document document)
            {
                if (document == null || !document.IsValidObject || document.IsFamilyDocument || document.IsLinked)
                    return new[] { string.Empty, string.Empty };
                string centralPath = string.Empty;
                if (document.IsWorkshared)
                {
                    try
                    {
                        using (ModelPath modelPath = document.GetWorksharingCentralModelPath())
                            if (modelPath != null)
                                centralPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath);
                    }
                    catch (Autodesk.Revit.Exceptions.InvalidOperationException) { }
                }
                return new[] { centralPath, document.PathName ?? string.Empty };
            }

            private ProjectMatch MatchModel(string[] paths)
            {
                // Центральный путь приоритетен. К локальному переходим только при отсутствии совпадений.
                var central = ProjectPathMatcher.Find(_projects, paths[0]);
                return central.Project != null || central.IsAmbiguous
                    ? central : ProjectPathMatcher.Find(_projects, paths[1]);
            }

            private void UpdateActiveProject(Document document, string filePath, bool force)
            {
                string[] paths = GetModelPaths(document);
                string displayPath = !string.IsNullOrWhiteSpace(filePath) ? filePath
                    : !string.IsNullOrWhiteSpace(paths[0]) ? paths[0]
                    : document == null || !document.IsValidObject ? "Нет открытого файла" : "Файл не сохранён";
                string key = displayPath + "\n" + string.Join("\n", paths);
                if (!force && string.Equals(_lastModelKey, key, StringComparison.Ordinal)) return;
                _lastModelKey = key;
                var match = MatchModel(paths);
                _owner.SetProjects(_projects, match, displayPath);
            }

            private static IList<OpenDocumentItem> GetOpenProjects(UIApplication app)
            {
                var items = new List<OpenDocumentItem>();
                foreach (Document doc in app.Application.Documents)
                {
                    if (!doc.IsValidObject || doc.IsFamilyDocument || doc.IsLinked || doc.IsReadOnly)
                        continue;
                    items.Add(new OpenDocumentItem
                    {
                        Document = doc,
                        DisplayName = doc.Title,
                        LoadTargetName = doc.Title + (string.IsNullOrWhiteSpace(doc.PathName)
                            ? " — не сохранён" : " — " + doc.PathName)
                    });
                }
                return items.OrderBy(p => p.DisplayName).ToList();
            }

            public void Execute(UIApplication app)
            {
                var request = PendingRequest;
                PendingRequest = null;
                string savedPath = null;
                try
                {
                    if (request == null) return;
                    string typeName = ValidateTypeName(request.TypeName);
                    Document target = null;
                    if (request.Kind == RequestKind.AddToProject)
                    {
                        // Список базы отражает активный проект; цель загрузки выбирается отдельно.
                        // Документы получаем заново непосредственно перед отдельным выбором цели.
                        var projects = GetOpenProjects(app);
                        if (projects.Count == 0)
                            throw new InvalidOperationException("Нет доступных проектов. Откройте проект в Revit и повторите загрузку.");
                        var project = _owner.ChooseLoadProject(projects,
                            app.ActiveUIDocument == null ? null : app.ActiveUIDocument.Document);
                        if (project == null)
                        {
                            _owner.SetStatus("Выбор проекта отменён. Файлы не изменены.", false);
                            return;
                        }
                        target = GetTargetDocument(app, project);
                    }

                    string sourcePath = ResolveSourcePath();
                    if (!File.Exists(sourcePath))
                        throw new FileNotFoundException("Исходное семейство недоступно. Проверьте подключение диска X: и доступ к файлу.\n" + sourcePath);

                    Document namingDocument = target ?? (app.ActiveUIDocument == null ? null : app.ActiveUIDocument.Document);
                    ProjectItem namingProject = MatchModel(GetModelPaths(namingDocument)).Project;
                    string outputPath = _owner.ChooseOutputPath(namingProject == null ? null : namingProject.Code);
                    if (outputPath == null)
                    {
                        _owner.SetStatus("Сохранение отменено. Файлы не изменены.", false);
                        return;
                    }
                    outputPath = ValidateOutputPath(app, sourcePath, outputPath);
                    _owner.SetStatus("Создание копии и переименование настроенного типа…", false);
                    string cleanupWarning = BuildAndSaveFamily(app, sourcePath, outputPath, typeName);
                    savedPath = outputPath;
                    _owner.RememberSavedPath(savedPath);

                    if (request.Kind == RequestKind.OpenInRevit)
                    {
                        // Документ копии уже закрыт; открытых транзакций здесь нет.
                        app.OpenAndActivateDocument(savedPath);
                        _owner.SetStatus("Вент. установка открыта с типом «" + typeName + "»." + cleanupWarning, false);
                    }
                    else
                    {
                        string result = LoadType(target, savedPath, typeName);
                        _owner.SetStatus(result + "\nКопия сохранена: " + savedPath + cleanupWarning, false);
                    }
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    _owner.SetStatus(savedPath == null ? "Операция отменена." :
                        "Открытие или загрузка отменены. Копия уже сохранена:\n" + savedPath, false);
                }
                catch (Exception ex)
                {
                    string prefix = savedPath == null ? "Операция не выполнена.\n" :
                        "Копия сохранена, но следующее действие не выполнено:\n" + savedPath + "\n";
                    _owner.SetStatus(prefix + ex.Message, true);
                }
                finally
                {
                    _owner.SetBusy(false);
                }
            }

            internal static string ValidateTypeName(string value)
            {
                string name = (value ?? string.Empty).Trim();
                if (name.Length == 0)
                    throw new ArgumentException("Введите имя нового типа.");
                if (string.Equals(name, BaseFamilyTypeName, StringComparison.OrdinalIgnoreCase))
                    throw new ArgumentException("Это имя базового типа. Введите другое имя для новой установки.");
                if (name.Any(char.IsControl) || name.IndexOfAny(new[] { '\\', ':', '{', '}', '[', ']', '|', ';', '<', '>', '?', '`', '~' }) >= 0)
                    throw new ArgumentException("Имя типа содержит недопустимые символы: \\ : { } [ ] | ; < > ? ` ~ или перевод строки.");
                return name;
            }

            private static Document GetTargetDocument(UIApplication app, OpenDocumentItem project)
            {
                Document target = project == null ? null : project.Document;
                if (target == null || !target.IsValidObject ||
                    !app.Application.Documents.Cast<Document>().Any(d => d.Equals(target)))
                    throw new InvalidOperationException("Откройте нужный проект в Revit и повторите действие.");
                if (target.IsFamilyDocument || target.IsLinked || target.IsReadOnly || target.IsModifiable)
                    throw new InvalidOperationException("Выбранный проект сейчас недоступен для загрузки. Завершите редактирование и повторите действие.");
                return target;
            }

            private static string ValidateOutputPath(UIApplication app, string sourcePath, string outputPath)
            {
                string path = Path.GetFullPath(outputPath);
                if (!string.Equals(Path.GetExtension(path), ".rfa", StringComparison.OrdinalIgnoreCase))
                    throw new ArgumentException("Для копии необходимо выбрать файл с расширением .rfa.");
                if (SamePath(path, sourcePath) || SamePath(path, SourceFamilyPath) || SamePath(path, LiteralSourceFamilyPath))
                    throw new InvalidOperationException("Нельзя сохранять копию поверх исходного семейства. Выберите другой файл.");
                // Другое имя семейства также защищает оригинал при доступе через UNC / сетевой диск
                // и исключает загрузку копии под исходным именем в проект.
                if (string.Equals(Path.GetFileName(path), Path.GetFileName(sourcePath), StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Задайте копии другое имя файла, чтобы отличать её от исходного семейства.");
                if (!Directory.Exists(Path.GetDirectoryName(path)))
                    throw new DirectoryNotFoundException("Выбранная папка недоступна.");
                foreach (Document doc in app.Application.Documents)
                {
                    if (doc.IsValidObject && !string.IsNullOrWhiteSpace(doc.PathName) && SamePath(doc.PathName, path))
                        throw new InvalidOperationException("Этот файл уже открыт в Revit. Закройте его с сохранением нужных изменений или выберите другое имя копии.");
                }
                if (File.Exists(path) && (File.GetAttributes(path) & FileAttributes.ReadOnly) != 0)
                    throw new IOException("Файл копии доступен только для чтения. Выберите другое место сохранения.");
                return path;
            }

            private static bool SamePath(string a, string b)
            {
                // Открытые облачные проекты могут иметь PathName вида BIM 360://... .
                // Такие адреса нельзя передавать в Path.GetFullPath как пути Windows.
                if (!Path.IsPathRooted(a) || !Path.IsPathRooted(b)) return false;
                return string.Equals(Path.GetFullPath(a), Path.GetFullPath(b), StringComparison.OrdinalIgnoreCase);
            }

            private static string BuildAndSaveFamily(UIApplication app, string sourcePath, string outputPath, string typeName)
            {
                // Подготовка на том же томе, что и результат: готовый RFA заменяет старый файл
                // только после успешной транзакции, SaveAs и Close. Исходник лишь читается File.Copy.
                string workDirectory = Path.Combine(Path.GetDirectoryName(outputPath),
                    ".KPLN_Ventilation_" + Guid.NewGuid().ToString("N"));
                string inputDirectory = Path.Combine(workDirectory, "input");
                string stagedDirectory = Path.Combine(workDirectory, "result");
                Document familyDoc = null;
                string warning = string.Empty;
                try
                {
                    Directory.CreateDirectory(inputDirectory);
                    Directory.CreateDirectory(stagedDirectory);
                    string inputPath = Path.Combine(inputDirectory, Path.GetFileName(sourcePath));
                    string stagedPath = Path.Combine(stagedDirectory, Path.GetFileName(outputPath));
                    File.Copy(sourcePath, inputPath, false);
                    File.SetAttributes(inputPath, File.GetAttributes(inputPath) & ~FileAttributes.ReadOnly);
                    familyDoc = app.Application.OpenDocumentFile(inputPath);
                    if (!familyDoc.IsFamilyDocument)
                        throw new InvalidOperationException("Исходный файл не является редактируемым семейством Revit.");

                    using (var transaction = new Transaction(familyDoc, "Переименовать тип вентиляционной установки"))
                    {
                        if (transaction.Start() != TransactionStatus.Started)
                            throw new InvalidOperationException("Не удалось начать изменение копии семейства.");
                        // Не оставляем Pending-транзакцию после возврата из обработчика.
                        var options = transaction.GetFailureHandlingOptions();
                        options.SetForcedModalHandling(true);
                        transaction.SetFailureHandlingOptions(options);
                        FamilyManager manager = familyDoc.FamilyManager;
                        RenameConfiguredType(manager, typeName);
                        ApplyConfiguredParameters(manager);
                        if (transaction.Commit() != TransactionStatus.Committed)
                            throw new InvalidOperationException("Revit отменил переименование типа. Файл копии не перезаписан.");
                    }

                    using (var options = new SaveAsOptions())
                    {
                        options.OverwriteExistingFile = false;
                        options.MaximumBackups = 1;
                        familyDoc.SaveAs(stagedPath, options);
                    }
                    if (!familyDoc.Close(false))
                        throw new InvalidOperationException("Не удалось закрыть подготовленную копию. Файл результата не перезаписан.");
                    familyDoc = null;

                    if (File.Exists(outputPath))
                    {
                        // Без небезопасного fallback Delete + Copy: если том не поддерживает Replace,
                        // операция завершится ошибкой, а существующий результат останется на месте.
                        File.Replace(stagedPath, outputPath, Path.Combine(workDirectory, "previous.rfa"));
                    }
                    else
                        File.Move(stagedPath, outputPath);
                }
                finally
                {
                    bool canClean = true;
                    if (familyDoc != null && familyDoc.IsValidObject)
                    {
                        try { canClean = familyDoc.Close(false); }
                        catch { canClean = false; }
                    }
                    if (canClean && Directory.Exists(workDirectory))
                    {
                        try { Directory.Delete(workDirectory, true); }
                        catch (IOException) { warning = "\nНе удалось удалить временную папку: " + workDirectory; }
                        catch (UnauthorizedAccessException) { warning = "\nНе удалось удалить временную папку: " + workDirectory; }
                    }
                    else if (!canClean)
                        warning = "\nВременный документ не закрыт: " + workDirectory;
                }
                return warning;
            }

            private static void RenameConfiguredType(FamilyManager manager, string typeName)
            {
                var types = manager.Types.Cast<FamilyType>().ToList();
                if (!types.Any(t => string.Equals(t.Name, BaseFamilyTypeName, StringComparison.Ordinal)))
                    throw new InvalidOperationException("Не найден базовый тип «" + BaseFamilyTypeName
                        + "». Копия не сохранена. Проверьте имя типа в исходном семействе.");

                var configuredType = types.SingleOrDefault(t =>
                    string.Equals(t.Name, ConfiguredFamilyTypeName, StringComparison.Ordinal));
                if (configuredType == null)
                    throw new InvalidOperationException("Не найден настроенный тип «" + ConfiguredFamilyTypeName
                        + "». Копия не сохранена. Проверьте состав типов исходного семейства.");

                if (types.Any(t => !string.Equals(t.Name, ConfiguredFamilyTypeName, StringComparison.Ordinal)
                    && string.Equals(t.Name, typeName, StringComparison.OrdinalIgnoreCase)))
                    throw new InvalidOperationException("Имя «" + typeName
                        + "» уже занято другим типом. Введите другое имя.");

                // Выбираем именно настроенную установку, независимо от текущего типа исходника.
                // Существующие типы и вложенные экземпляры не удаляем, новый тип не создаём.
                manager.CurrentType = configuredType;
                if (!string.Equals(configuredType.Name, typeName, StringComparison.Ordinal))
                    manager.RenameCurrentType(typeName);
            }

            private static void ApplyConfiguredParameters(FamilyManager manager)
            {
                // Точка расширения следующего этапа. Уже выбран и переименован настроенный тип,
                // транзакция открыта. Здесь будут проверка и manager.Set(...) для согласованных
                // параметров; преобразования единиц и различия API — только через IDHelper.
                // Пока значения параметров настроенного типа сохраняются без изменений.
            }

            private static string LoadType(Document target, string path, string typeName)
            {
                var loadOptions = new FamilyLoadOptions(Path.GetFileNameWithoutExtension(path));
                using (var transaction = new Transaction(target, "Загрузить тип вентиляционной установки"))
                {
                    if (transaction.Start() != TransactionStatus.Started)
                        throw new InvalidOperationException("Не удалось начать загрузку в проект.");
                    var failures = transaction.GetFailureHandlingOptions();
                    failures.SetForcedModalHandling(true);
                    transaction.SetFailureHandlingOptions(failures);
                    FamilySymbol symbol;
                    bool loaded = target.LoadFamilySymbol(path, typeName, loadOptions, out symbol);
                    if (loadOptions.Cancelled)
                    {
                        transaction.RollBack();
                        return "Загрузка в проект отменена.";
                    }
                    if (!loaded || symbol == null)
                    {
                        transaction.RollBack();
                        return "Revit не выполнил загрузку типа в проект. Возможно, такая версия типа уже загружена.";
                    }
                    if (!symbol.IsActive) symbol.Activate();
                    if (transaction.Commit() != TransactionStatus.Committed)
                        throw new InvalidOperationException("Revit отменил загрузку типа в проект.");
                }
                return "Тип «" + typeName + "» загружен в проект «" + target.Title + "». Экземпляр не размещён.";
            }
        }

        private sealed class FamilyLoadOptions : IFamilyLoadOptions
        {
            private readonly string _rootFamilyName;
            private bool? _overwriteAccepted;
            internal bool Cancelled { get; private set; }

            internal FamilyLoadOptions(string rootFamilyName) { _rootFamilyName = rootFamilyName; }

            public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
            {
                bool accepted = ConfirmOverwrite();
                overwriteParameterValues = accepted;
                return accepted;
            }

            private bool ConfirmOverwrite()
            {
                if (_overwriteAccepted.HasValue) return _overwriteAccepted.Value;
                var dialog = new TaskDialog(PluginName)
                {
                    MainInstruction = "В проекте уже есть это семейство. Обновить его?",
                    MainContent = "Будут обновлены определение семейства и значения типов из загружаемого семейства. Это может изменить уже размещённые экземпляры.",
                    CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No,
                    DefaultButton = TaskDialogResult.No
                };
                bool accepted = dialog.Show() == TaskDialogResult.Yes;
                _overwriteAccepted = accepted;
                Cancelled = !accepted;
                return accepted;
            }

            public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse,
                out FamilySource source, out bool overwriteParameterValues)
            {
                // Основное семейство само тоже может быть общим: для него разрешаем
                // обновление только после того же подтверждения, что и для обычного.
                if (string.Equals(sharedFamily.Name, _rootFamilyName, StringComparison.OrdinalIgnoreCase))
                {
                    bool accepted = ConfirmOverwrite();
                    source = accepted ? FamilySource.Family : FamilySource.Project;
                    overwriteParameterValues = accepted;
                    return accepted;
                }
                // Существующие общие вложенные семейства проекта сохраняем.
                source = FamilySource.Project;
                overwriteParameterValues = false;
                return true;
            }
        }
    }
}
