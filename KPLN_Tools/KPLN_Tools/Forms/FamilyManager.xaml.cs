using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Action = System.Action;
using Application = System.Windows.Application;
using Button = System.Windows.Controls.Button;
using ComboBox = System.Windows.Controls.ComboBox;
using Grid = System.Windows.Controls.Grid;
using TextBox = System.Windows.Controls.TextBox;
using Window = System.Windows.Window;


namespace KPLN_Tools.Forms
{
    // ExternalEventsHost
    internal static class ExternalEventsHost
    {
        public static ExternalEvent LoadFamilyEvent;
        public static LoadFamilyHandler LoadFamilyHandler;

        public static ExternalEvent BulkPagedUpdateEvent;
        public static BulkPagedUpdateHandler BulkPagedUpdateHandler;

        public static void EnsureCreated()
        {
            if (LoadFamilyEvent == null)
            {
                LoadFamilyHandler = new LoadFamilyHandler();
                LoadFamilyEvent = ExternalEvent.Create(LoadFamilyHandler);
            }

            if (BulkPagedUpdateEvent == null)
            {
                BulkPagedUpdateHandler = new BulkPagedUpdateHandler();
                BulkPagedUpdateEvent = ExternalEvent.Create(BulkPagedUpdateHandler);
            }
        }
    }

    internal class LoadFamilyHandler : IExternalEventHandler
    {
        public string FilePath;
        public List<string> BatchPaths;

        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            if (uidoc == null)
            {
                TaskDialog.Show("Ошибка", "Нет активного документа.");
                return;
            }

            Document targetDoc = uidoc.Document;
            if (targetDoc.IsFamilyDocument)
            {
                targetDoc = app.Application.Documents
                    .Cast<Document>()
                    .FirstOrDefault(d => !d.IsFamilyDocument);

                if (targetDoc == null)
                {
                    TaskDialog.Show("Ошибка", "Нет открытого проекта для загрузки семейста или открыт файл семейства.");
                    return;
                }
            }

            var paths = (BatchPaths != null && BatchPaths.Count > 0)
                ? BatchPaths
                : (string.IsNullOrWhiteSpace(FilePath) ? new List<string>() : new List<string> { FilePath });

            paths = paths
                .Where(p => !string.IsNullOrWhiteSpace(p) && File.Exists(p))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (paths.Count == 0)
            {
                TaskDialog.Show("Загрузка семейств", "Нет валидных путей для загрузки.");
                BatchPaths = null; FilePath = null;
                return;
            }

            bool single = paths.Count == 1;

            int loaded = 0;
            var failedNames = new List<string>();
            var opts = new ReloadFamilyLoadOptions();

            using (var t = new Transaction(targetDoc, "KPLN. Загрузить семейства"))
            {
                t.Start();

                var fho = t.GetFailureHandlingOptions();
                fho.SetFailuresPreprocessor(new SilentFailuresPreprocessor());
                t.SetFailureHandlingOptions(fho);
                foreach (var p in paths)
                {
                    string nameForReport = System.IO.Path.GetFileNameWithoutExtension(p);

                    try
                    {
                        var optsLocal = new ReloadFamilyLoadOptions();

                        if (targetDoc.LoadFamily(p, optsLocal, out Family famLoaded))
                        {
                            loaded++;
                            continue;
                        }

                        string internalName = GetFamilyInternalName(app, p) ?? nameForReport;

                        var existing = new FilteredElementCollector(targetDoc)
                            .OfClass(typeof(Family))
                            .Cast<Family>()
                            .FirstOrDefault(f => string.Equals(f.Name, internalName, StringComparison.OrdinalIgnoreCase));

                        if (existing == null)
                        {
                            failedNames.Add(nameForReport);
                            continue;
                        }

                        if (FamilyHasInstances(targetDoc, existing))
                        {
                            failedNames.Add(nameForReport);
                            continue;
                        }

                        try
                        {
                            targetDoc.Delete(existing.Id);

                            var optsAfterDelete = new ReloadFamilyLoadOptions();
                            if (targetDoc.LoadFamily(p, optsAfterDelete, out Family famAfter))
                                loaded++;
                            else
                                failedNames.Add(nameForReport);
                        }
                        catch
                        {
                            failedNames.Add(nameForReport);
                        }
                    }
                    catch
                    {
                        failedNames.Add(nameForReport);
                    }
                }

                t.Commit();
            }

            if (single)
            {
                TaskDialog.Show("Загрузка семейства", loaded == 1 ? "Семейство добавлено/обновлено в проекте." : "ОШИБКА. Данное семейство не было добавлено или обновлено в проекте.");
            }
            else
            {
                if (loaded == paths.Count)
                {
                    TaskDialog.Show("Загрузка семейств", "Все семейства добавлены проект.");
                }
                else
                {
                    var sb = new StringBuilder();
                    sb.AppendLine($"Загружено/Обновлено: {loaded} из {paths.Count}");
                    if (failedNames.Count > 0)
                    {
                        sb.AppendLine("Не добавлены:");
                        foreach (var n in failedNames) 
                            sb.AppendLine(EllipsisEnd(n, 37));
                    }
                    TaskDialog.Show("Загрузка семейств", sb.ToString());
                }
            }

            BatchPaths = null;
            FilePath = null;
        }

        public string GetName() => "KPLN.LoadFamilyHandler";

        private static string GetFamilyInternalName(UIApplication uiapp, string rfaPath)
        {
            try
            {
                var famDoc = uiapp.Application.OpenDocumentFile(rfaPath);
                if (famDoc == null || !famDoc.IsFamilyDocument)
                    return null;

                string name = famDoc.OwnerFamily?.Name;
                try { famDoc.Close(false); } catch { }
                return string.IsNullOrWhiteSpace(name) ? null : name;
            }
            catch
            {
                return null;
            }
        }

        private static bool FamilyHasInstances(Document doc, Family family)
        {
            if (family == null) return false;

            var instUsed = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Any(fi => fi.Symbol != null && fi.Symbol.Family != null && fi.Symbol.Family.Id == family.Id);

            return instUsed;
        }

        private static string EllipsisEnd(string s, int max = 40)
        {
            if (string.IsNullOrEmpty(s) || max <= 0) return "";
            if (s.Length <= max) return s;
            return s.Substring(0, Math.Max(1, max - 1)) + "…";
        }
    }

    internal class SilentFailuresPreprocessor : IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(FailuresAccessor a)
        {
            var msgs = a.GetFailureMessages();
            if (msgs != null)
                foreach (var m in msgs)
                    if (m.GetSeverity() == FailureSeverity.Warning)
                        a.DeleteWarning(m);
            return FailureProcessingResult.Continue;
        }
    }

    internal class ReloadFamilyLoadOptions : IFamilyLoadOptions
    {
        private const bool OVERWRITE_PARAMS = true; 

        public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
        {
            overwriteParameterValues = OVERWRITE_PARAMS;
            return true; 
        }

        public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse,
            out FamilySource source, out bool overwriteParameterValues)
        {
            source = FamilySource.Family; 
            overwriteParameterValues = OVERWRITE_PARAMS;
            return true; 
        }
    }

    // IExternalEventHandler. Загрузка данных из семейств
    internal class BulkPagedUpdateHandler : IExternalEventHandler
    {
        public string DbPath;
        public int PageSize = 200;
        public volatile bool IsRunning;
        public int Selected, Updated, Skipped, Errors;
        public Action<BulkPagedUpdateHandler> Completed;

        public class ErrorEntry
        {
            public int Id { get; set; }
            public string FullPath { get; set; }
            public string Message { get; set; }
        }
        public List<ErrorEntry> ErrorLog { get; } = new List<ErrorEntry>();

        private void AddError(int id, string fullPath, string message)
        {
            Errors++;
            try
            {
                ErrorLog.Add(new ErrorEntry { Id = id, FullPath = fullPath, Message = message });
            }
            catch {}
        }

        public void Execute(UIApplication app)
        {
            IsRunning = true;
            Selected = Updated = Skipped = Errors = 0;

            var detach = AttachRevitDialogSuppressors(app);

            try
            {
                using (var conn = new SQLiteConnection($"Data Source={DbPath};Version=3;"))
                {
                    conn.Open();

                    long lastId = 0;
                    while (true)
                    {
                        var ids = new List<int>();
                        using (var cmd = new SQLiteCommand(@"
                        SELECT ID
                        FROM FamilyManager
                        WHERE STATUS NOT IN ('ABSENT','ERROR','IGNORE')
                          AND (IMPORT_INFO IS NULL OR length(IMPORT_INFO)=0)
                          AND ID > @last
                        ORDER BY ID
                        LIMIT @lim;", conn))
                        {
                            cmd.Parameters.AddWithValue("@last", lastId);
                            cmd.Parameters.AddWithValue("@lim", PageSize);
                            using (var rd = cmd.ExecuteReader())
                                while (rd.Read()) ids.Add(Convert.ToInt32(rd["ID"]));
                        }

                        if (ids.Count == 0) break;

                        Selected += ids.Count;

                        using (var tx = conn.BeginTransaction())
                        using (var cmdRead = new SQLiteCommand(
                            @"SELECT STATUS, FULLPATH, DEPARTAMENT FROM FamilyManager WHERE ID=@id LIMIT 1;", conn, tx))
                        using (var cmdUpd = new SQLiteCommand(
                            @"UPDATE FamilyManager SET IMPORT_INFO=@json WHERE ID=@id;", conn, tx))
                        {
                            var pIdRead = cmdRead.Parameters.Add("@id", System.Data.DbType.Int32);
                            var pIdUpd = cmdUpd.Parameters.Add("@id", System.Data.DbType.Int32);
                            var pJsonUpd = cmdUpd.Parameters.Add("@json", System.Data.DbType.String);

                            foreach (int id in ids)
                            {
                                lastId = id;

                                string status = null;
                                string full = null; 
                                string dep = null;

                                try
                                {
                                    pIdRead.Value = id;
                                    using (var rd = cmdRead.ExecuteReader(System.Data.CommandBehavior.SingleRow))
                                    {
                                        if (!rd.Read()) { Skipped++; continue; }
                                        status = rd["STATUS"] as string;
                                        full = rd["FULLPATH"] as string;
                                        dep = rd["DEPARTAMENT"] as string;
                                    }

                                    if (string.Equals(status, "ABSENT", StringComparison.OrdinalIgnoreCase) ||
                                        string.Equals(status, "ERROR", StringComparison.OrdinalIgnoreCase) ||
                                        string.Equals(status, "IGNORE", StringComparison.OrdinalIgnoreCase))
                                    { 
                                        Skipped++;
                                        continue; 
                                    }

                                    if (string.IsNullOrWhiteSpace(dep)) 
                                    { 
                                        Skipped++; 
                                        continue; 
                                    }
                                    if (string.IsNullOrWhiteSpace(full) || !File.Exists(full))
                                    { 
                                        Errors++;
                                        AddError(id, full, "Файл семейства не найден по пути FULLPATH.");
                                        continue; 
                                    }

                                    string json = FamilyManager.ReadImportInfoFromFamily(app, full, dep);
                                    if (string.IsNullOrWhiteSpace(json)) 
                                    { 
                                        Skipped++; 
                                        continue; 
                                    }

                                    pIdUpd.Value = id;
                                    pJsonUpd.Value = json;
                                    int rows = cmdUpd.ExecuteNonQuery();

                                    if (rows > 0)
                                    {
                                        Updated++;
                                    }
                                    else
                                    {
                                        Errors++;
                                        AddError(id, full, "UPDATE FamilyManager не изменил запись (rows=0).");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Errors++;
                                    AddError(id, full, ex.Message);
                                }
                            }

                            tx.Commit();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                AddError(0, null, "Глобальная ошибка обработки: " + ex.Message);
                Errors++;
            }
            finally
            {
                try { detach(); } catch { }
                IsRunning = false;

                if (Completed != null)
                {
                    EventHandler<IdlingEventArgs> idl = null;
                    idl = (s2, e2) =>
                    {
                        try { app.Idling -= idl; } catch { }
                        try { Completed?.Invoke(this); } finally { Completed = null; }
                    };
                    app.Idling += idl;
                }
            }
        }

        public string GetName() => "KPLN.BulkPagedUpdateHandler";

        private static Action AttachRevitDialogSuppressors(UIApplication uiapp)
        {
            EventHandler<DialogBoxShowingEventArgs> dbHandler = (s, e) =>
            {
                try
                {
                    if (e is TaskDialogShowingEventArgs td) td.OverrideResult((int)TaskDialogResult.Ok);
                    else if (e is MessageBoxShowingEventArgs mb) mb.OverrideResult((int)System.Windows.Forms.DialogResult.OK);
                    else e.OverrideResult(1);
                }
                catch { }
            };

            EventHandler<FailuresProcessingEventArgs> fpHandler = (s, args) =>
            {
                try
                {
                    var fa = args.GetFailuresAccessor();
                    foreach (var f in fa.GetFailureMessages())
                    {
                        if (f.GetSeverity() == FailureSeverity.Warning) fa.DeleteWarning(f);
                        else fa.ResolveFailure(f);
                    }
                    args.SetProcessingResult(FailureProcessingResult.Continue);
                }
                catch { }
            };

            uiapp.DialogBoxShowing += dbHandler;
            uiapp.Application.FailuresProcessing += fpHandler;

            return () =>
            {
                try { uiapp.DialogBoxShowing -= dbHandler; } catch { }
                try { uiapp.Application.FailuresProcessing -= fpHandler; } catch { }
            };
        }
    }

    // Данные из БД
    public class FamilyManagerRecord
    {
        public int ID { get; set; }
        public string Status { get; set; }
        public string FullPath { get; set; }
        public int Category { get; set; }
        public int SubCategory { get; set; }
        public int Project { get; set; }
        public int Stage { get; set; }
        public string Departament { get; set; }
        public string ImportInfo { get; set; }
        public byte[] ImageBytes { get; set; }
    }

    // Док-панель
    public partial class FamilyManager : UserControl
    {
        private UIApplication _uiapp;
        string _currentSubDep;
        private bool _suppressAutoReload = false;

        private const string ITEM_ERROR = "ОШИБКА";
        private const string DB_PATH = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_FamilyManager.db";
        private const string RFA_ROOT = @"X:\BIM\3_Семейства";
        private List<FamilyManagerRecord> _records;

        private const string BIM = "BIM";
        private const string BIM_ADMIN = "BIM (Админ)";
        private static bool IsBimAdmin(string dep)
            => string.Equals(dep?.Trim(), BIM_ADMIN, StringComparison.OrdinalIgnoreCase);
        private static string DepForDb(string dep)
            => IsBimAdmin(dep) ? BIM : (dep ?? "").Trim();
     
        private StackPanel _bimRootPanel;
        private Dictionary<string, bool> _bimExpandState = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        private double _bimScrollOffset = 0;
        private ScrollViewer _bimScrollViewer;

        private StackPanel _universalRootPanel;
        private string _universalDepartment;
        private readonly Dictionary<string, bool> _univExpandState = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        private ScrollViewer _univScrollViewer;
        private double _univScrollOffset = 0;
        private ComboBox _cbStage;
        private ComboBox _cbProject;

        private readonly Dictionary<string, BitmapImage> _embeddedIconCache = new Dictionary<string, BitmapImage>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<(int CatId, int SubId), BitmapImage> _categoryIconCache = new Dictionary<(int, int), BitmapImage>();
        private readonly Dictionary<(int CatId, int SubId), byte[]> _categoryPics = new Dictionary<(int, int), byte[]>();

        private readonly HashSet<int> _favoriteIds = new HashSet<int>();
        private string _favoritesPath;
        private DateTime _favoritesLastWrite = DateTime.MinValue;
        private FileSystemWatcher _favoritesWatcher;
        private bool IsFavorite(FamilyManagerRecord r) => r != null && _favoriteIds.Contains(r.ID);

        private readonly Dictionary<(int Id, string Name), Dictionary<int, string>> _categoriesById = new Dictionary<(int Id, string Name), Dictionary<int, string>>();

        private Dictionary<int, string> _stagesById = new Dictionary<int, string>();
        private Dictionary<string, int> _stageIdByName = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        private Dictionary<int, string> _projectsById = new Dictionary<int, string>();
        private Dictionary<string, int> _projectIdByName = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        private readonly Dictionary<string, HashSet<int>> _selectedByDept  = new Dictionary<string, HashSet<int>>(StringComparer.OrdinalIgnoreCase);
        private string GetDeptKey() => DepForDb(_universalDepartment ?? GetCurrentDepartment()) ?? "";

        private FrameworkElement _scenarioContent;

        private TextBox _tbSearch;
        private const string SEARCH_WATERMARK = "Поиск по названию и параметрам";
        private bool _isWatermarkActive = false;
        private bool _isApplyingWatermark = false;
        private bool _depsTried = false;
        private bool _depsLoaded = false;
        private Dictionary<int, string> _searchIndex = new Dictionary<int, string>();
        private DispatcherTimer _searchDebounceTimer;

        private Button _btnOpenInRevit;   
        private Button _btnLoadIntoProject;

        public void SetUIApplication(UIApplication uiapp)
        {
            _uiapp = uiapp;
            ResetDeptCacheAndUi();
        }

        public FamilyManager(string currentStr)
        {
            InitializeComponent();
            _currentSubDep = currentStr;
        }

        // Получение названия отдела из XAML
        private string GetCurrentDepartment()
        {
            return CmbDepartment?.SelectedItem?.ToString()?.Trim();
        }

        // Данные БД. Отделы
        private static List<string> LoadDepartments(string dbPath)
        {
            var result = new List<string>();
            if (!File.Exists(dbPath)) return result;

            string connStr = $"Data Source={dbPath};Version=3;";
            using (var conn = new SQLiteConnection(connStr))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand(
                    "SELECT DISTINCT TRIM(NAME) " +
                    "FROM Departament " +
                    "WHERE NAME IS NOT NULL AND TRIM(NAME) <> '' " +
                    "ORDER BY ID;", conn))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var name = reader.GetString(0)?.Trim();
                        if (!string.IsNullOrWhiteSpace(name))
                            result.Add(name);
                    }
                }
            }
            return result;
        }

        // Список отделов
        private void EnsureDepartmentsLoadedIntoComboPreservingSelection()
        {
            try
            {
                var deps = LoadDepartments(DB_PATH);
                if (deps == null || deps.Count == 0)
                {
                    CmbDepartment.ItemsSource = new List<string> { ITEM_ERROR };
                    CmbDepartment.SelectedItem = ITEM_ERROR;
                    _depsLoaded = false;
                    UpdateUiState();
                    return;
                }

                var withUi = new List<string> { BIM_ADMIN, BIM };
                foreach (var d in deps)
                    if (!withUi.Contains(d, StringComparer.OrdinalIgnoreCase))
                        withUi.Add(d);

                var prev = CmbDepartment.SelectedItem?.ToString();

                CmbDepartment.ItemsSource = withUi;

                if (!string.IsNullOrWhiteSpace(prev) &&
                    withUi.Any(x => x.Equals(prev, StringComparison.OrdinalIgnoreCase)))
                {
                    CmbDepartment.SelectedItem = withUi.First(x => x.Equals(prev, StringComparison.OrdinalIgnoreCase));
                }
                else
                {
                    CmbDepartment.SelectedItem = withUi.First(); // по умолчанию BIM (Админ)
                }

                _depsLoaded = true;
                UpdateUiState();
            }
            catch
            {
                CmbDepartment.ItemsSource = new List<string> { ITEM_ERROR };
                CmbDepartment.SelectedItem = ITEM_ERROR;
                _depsLoaded = false;
                UpdateUiState();
            }
        }

        // Список отделов. Сброс состояния
        private void ResetDeptCacheAndUi()
        {
            _depsTried = false;    
            _depsLoaded = false;

            if (CmbDepartment != null)
                EnsureDepartmentsLoadedIntoComboPreservingSelection();
        }

        // Данные БД. Проверка того, что наш отдел содержится в БД
        private static bool DepartmentExistsInDb(string dbPath, string depName)
        {
            if (string.IsNullOrWhiteSpace(depName)) return false;
            if (!File.Exists(dbPath)) return false;

            string connStr = $"Data Source={dbPath};Version=3;";
            using (var conn = new SQLiteConnection(connStr))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand(
                    "SELECT 1 FROM Departament WHERE TRIM(NAME) = TRIM(@name) COLLATE NOCASE LIMIT 1;", conn))
                {
                    cmd.Parameters.AddWithValue("@name", depName);
                    var obj = cmd.ExecuteScalar();
                    return obj != null;
                }
            }
        }

        // Обновление статуса (доступности кнопок)
        private void UpdateUiState()
        {
            string dep = CmbDepartment.SelectedItem?.ToString();
            bool isError = dep == ITEM_ERROR;

            CmbDepartment.IsEnabled = string.Equals(_currentSubDep, BIM, StringComparison.OrdinalIgnoreCase);
            BtnReload.IsEnabled = !isError;
            BtnSettings.IsEnabled = IsBimAdmin(dep); 
        }

        // Данные БД. Семейства.
        private static List<FamilyManagerRecord> LoadFamilyManagerRecords(string dbPath)
        {
            if (!File.Exists(dbPath))
            {
                TaskDialog.Show("Ошибка", "Файл БД не найден.");
                return null;
            }

            var result = new List<FamilyManagerRecord>();
            string connStr = $"Data Source={dbPath};Version=3;";

            using (var conn = new SQLiteConnection(connStr))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand(
                    "SELECT ID, STATUS, FULLPATH, LM_DATE, CATEGORY, PROJECT, STAGE, DEPARTAMENT, IMPORT_INFO FROM FamilyManager",
                    conn))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var rec = new FamilyManagerRecord
                        {
                            ID = reader.GetInt32(0),
                            Status = reader.GetString(1),
                            FullPath = reader.GetString(2),
                            Category = reader.GetInt32(4),
                            Project = reader.GetInt32(5),
                            Stage = reader.GetInt32(6),
                            Departament = reader.IsDBNull(7) ? null : reader.GetString(7),
                            ImportInfo = reader.IsDBNull(8) ? null : reader.GetString(8),
                        };
                        result.Add(rec);
                    }
                }
            }
            return result;
        }

        // Данные БД. Семейства ( + фильтр по отделу)
        private static List<FamilyManagerRecord> LoadFamilyManagerRecordsForDepartment(string dbPath, string depUi)
        {
            var result = new List<FamilyManagerRecord>();
            if (!File.Exists(dbPath)) return result;

            string depDb = DepForDb(depUi);
            if (string.IsNullOrWhiteSpace(depDb)) return result;

            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand(@"
                    SELECT ID, STATUS, FULLPATH, CATEGORY, SUB_CATEGORY, PROJECT, STAGE, DEPARTAMENT, IMPORT_INFO, IMAGE
                    FROM FamilyManager
                    WHERE DEPARTAMENT IS NOT NULL
                      AND DEPARTAMENT LIKE '%' || @dep || '%' COLLATE NOCASE
                      AND STATUS IN ('NEW','OK');
                    ", conn))
                {
                    cmd.Parameters.AddWithValue("@dep", depDb);

                    using (var rd = cmd.ExecuteReader())
                    {
                        while (rd.Read())
                        {
                            var rec = new FamilyManagerRecord
                            {
                                ID = rd.GetInt32(0),
                                Status = rd.IsDBNull(1) ? null : rd.GetString(1),
                                FullPath = rd.IsDBNull(2) ? null : rd.GetString(2),
                                Category = rd.IsDBNull(3) ? 0 : rd.GetInt32(3),
                                SubCategory = rd.IsDBNull(4) ? 0 : rd.GetInt32(4),
                                Project = rd.IsDBNull(5) ? 0 : rd.GetInt32(5),
                                Stage = rd.IsDBNull(6) ? 0 : rd.GetInt32(6),
                                Departament = rd.IsDBNull(7) ? null : rd.GetString(7),
                                ImportInfo = rd.IsDBNull(8) ? null : rd.GetString(8),                              
                                ImageBytes = rd.IsDBNull(9) ? null : (byte[])rd[9],
                            };
                            result.Add(rec);
                        }
                    }
                }
            }
            return result;
        }

        // Данные БД. Категория, Проект и стадия
        private void EnsureLookupsLoaded()
        {
            if (_stagesById.Count == 0) LoadStages(DB_PATH);
            if (_projectsById.Count == 0) LoadProjects(DB_PATH);
            if (_categoriesById.Count == 0) LoadCategories(DB_PATH);
            if (_categoryPics.Count == 0) LoadCategoryPics(DB_PATH);
        }

        // Данные БД. Стадия
        private void LoadStages(string dbPath)
        {
            _stagesById.Clear();
            _stageIdByName.Clear();

            if (!File.Exists(dbPath)) return;

            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("SELECT ID, NAME FROM Stage ORDER BY ID;", conn))
                using (var rd = cmd.ExecuteReader())
                {
                    while (rd.Read())
                    {
                        int id = Convert.ToInt32(rd["ID"]);
                        string name = (rd["NAME"] as string)?.Trim();
                        if (string.IsNullOrWhiteSpace(name)) continue;

                        _stagesById[id] = name;
                        if (!_stageIdByName.ContainsKey(name)) _stageIdByName[name] = id;
                    }
                }
            }
        }

        // Данные БД. Категории
        private void LoadCategories(string dbPath)
        {
            _categoriesById.Clear();
            if (!File.Exists(dbPath)) return;

            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand(
                    @"SELECT ID, NAME, NC_NAME 
                      FROM Category
                      WHERE NC_NAME IS NOT NULL;", conn))
                using (var rd = cmd.ExecuteReader())
                {
                    while (rd.Read())
                    {
                        int categoryId = Convert.ToInt32(rd.GetValue(0));
                        string categoryName = Convert.ToString(rd.GetValue(1));
                        string ncNameRaw = Convert.ToString(rd.GetValue(2));

                        var subMap = ParseSubcategoriesJson(ncNameRaw);
                        _categoriesById[(categoryId, categoryName)] = subMap ?? new Dictionary<int, string>();
                    }
                }
            }
        }

        // Данные БД. Парсинг JSON
        private static Dictionary<int, string> ParseSubcategoriesJson(string json)
        {
            var result = new Dictionary<int, string>();
            if (string.IsNullOrWhiteSpace(json)) return result;

            string s = json.Trim();
            if (!s.StartsWith("[")) return result; 

            try
            {
                var arr = JArray.Parse(s);
                foreach (var obj in arr.OfType<JObject>())
                {
                    var idTok = obj["id"];
                    var nameTok = obj["name"];
                    if (idTok == null || nameTok == null) continue;

                    int subId = idTok.Type == JTokenType.Integer
                        ? idTok.Value<int>()
                        : Convert.ToInt32(idTok.ToString(), System.Globalization.CultureInfo.InvariantCulture);

                    string name = nameTok.ToString()?.Trim();
                    if (subId >= 0 && !string.IsNullOrWhiteSpace(name))
                        result[subId] = name;
                }
            }
            catch{}
            return result;
        }

        // Данные БД. Проект
        private void LoadProjects(string dbPath)
        {
            _projectsById.Clear();
            _projectIdByName.Clear();

            if (!File.Exists(dbPath)) return;

            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("SELECT ID, NAME FROM Project ORDER BY NAME;", conn))
                using (var rd = cmd.ExecuteReader())
                {
                    while (rd.Read())
                    {
                        int id = Convert.ToInt32(rd["ID"]);
                        string name = (rd["NAME"] as string)?.Trim();
                        if (string.IsNullOrWhiteSpace(name)) continue;

                        _projectsById[id] = name;
                        if (!_projectIdByName.ContainsKey(name)) _projectIdByName[name] = id;
                    }
                }
            }
        }

        // XAML. Док панель. Загрузка формы
        // Стартовый список отделов из конструктора (без БД)
        private void UserControl_Loaded(object sender, System.Windows.RoutedEventArgs e)
        {
            var items = new List<string>();
            if (!string.IsNullOrWhiteSpace(_currentSubDep))
            {
                if (string.Equals(_currentSubDep, BIM, StringComparison.OrdinalIgnoreCase))
                {
                    items.Add(BIM_ADMIN);
                    items.Add(BIM);
                }
                else
                {
                    items.Add(_currentSubDep);
                }
            }
            else
            {
                items.Add(ITEM_ERROR);
            }

            CmbDepartment.ItemsSource = items;
            CmbDepartment.SelectedItem = items.First();
            UpdateUiState();

            EnsureDepartmentsLoadedIntoComboPreservingSelection();
            LoadFavoritesFromFile();
            StartFavoritesWatcher();

            _suppressAutoReload = false; 
            ReloadData();
        }

        // XAML. Док панель. Обновление статуса CmbDepartment (
        private void CmbDepartment_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            ClearCurrentDeptSelection();
            MainArea.Child = null;

            var dep = CmbDepartment.SelectedItem?.ToString();

            _universalDepartment = IsBimAdmin(dep) ? null : dep;

            if (string.Equals(dep, BIM, StringComparison.OrdinalIgnoreCase) || IsBimAdmin(dep))
            {
                if (!_depsLoaded || (CmbDepartment?.Items?.Count ?? 0) <= 2)
                {
                    EnsureDepartmentsLoadedIntoComboPreservingSelection();
                }
            }

            UpdateUiState();

            if (!_suppressAutoReload)
                ReloadData();
        }

        // XAML. Док панель. Загрузка данных по кнопке
        private void BtnReload_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            ReloadData();
        }

        // Вспомогательный метод. Загрузка данных по кнопке
        private void ReloadData()
        {
            var depUi = GetCurrentDepartment();
            var depDb = DepForDb(depUi);

            LoadFavoritesFromFile();

            if (!_depsLoaded || (CmbDepartment?.Items?.Count ?? 0) <= 2)
                EnsureDepartmentsLoadedIntoComboPreservingSelection();

            if (DepartmentExistsInDb(DB_PATH, depDb))
            {
                try
                {
                    if (string.Equals(depUi, BIM_ADMIN, StringComparison.OrdinalIgnoreCase))
                        _records = LoadFamilyManagerRecords(DB_PATH);
                    else
                        _records = LoadFamilyManagerRecordsForDepartment(DB_PATH, depUi);

                    RebuildSearchIndex();
                    BuildMainArea(depUi);
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Ошибка чтения БД", $"{ex}");
                }
            }
            else
            {
                TaskDialog.Show("Ошибка", "Не найден соответствующий отдел в БД.");
            }
        }

        // Док панель. Поиск. Индексация IMPORT_INFO. Собираем строковые/числовые значения из JSON
        private static void CollectJsonValues(JToken token, List<string> bag)
        {
            if (token == null) return;

            switch (token.Type)
            {
                case JTokenType.String:
                    {
                        var s = token.Value<string>();
                        if (!string.IsNullOrWhiteSpace(s)) bag.Add(s);
                        break;
                    }
                case JTokenType.Integer:
                case JTokenType.Float:
                case JTokenType.Boolean:
                    bag.Add(token.ToString());
                    break;

                case JTokenType.Array:
                    foreach (var it in token.Children())
                        CollectJsonValues(it, bag);
                    break;

                case JTokenType.Object:
                    foreach (var prop in ((JObject)token).Properties())
                        CollectJsonValues(prop.Value, bag);
                    break;

                default:
                    break;
            }
        }

        // Док панель. Поиск. Индексация IMPORT_INFO. Строим текст для поиска: имя файла + все значения из IMPORT_INFO
        private static string BuildSearchText(FamilyManagerRecord r)
        {
            var sb = new StringBuilder(256);

            var name = SafeFileName(r.FullPath);
            if (!string.IsNullOrWhiteSpace(name))
                sb.Append(name).Append(' ');

            if (!string.IsNullOrWhiteSpace(r.ImportInfo))
            {
                try
                {
                    var root = JToken.Parse(r.ImportInfo);
                    var bag = new List<string>(8);
                    CollectJsonValues(root, bag);

                    foreach (var v in bag)
                    {
                        var t = (v ?? "").Trim();
                        if (t.Length > 0)
                            sb.Append(t).Append(' ');
                    }
                }
                catch
                {
                }
            }

            return sb.ToString().ToUpperInvariant();
        }

        // Док панель. Поиск. Индексация IMPORT_INFO. Перестроение индекса для всех записей
        private void RebuildSearchIndex()
        {
            _searchIndex.Clear();
            var src = _records ?? new List<FamilyManagerRecord>();
            foreach (var r in src)
                _searchIndex[r.ID] = BuildSearchText(r);
        }

        // Загрузка данных по кнопке. Построение интерфейса в MainArea
        private void BuildMainArea(string subDepartamentName)
        {
            var root = new Grid
            {
                Margin = new Thickness(8)
            };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            _tbSearch = new TextBox
            {
                MinWidth = 200,
                Height = 26,
                Margin = new Thickness(0, 0, 0, 8),
                Padding = new Thickness(6, 2, 6, 2),
                VerticalContentAlignment = VerticalAlignment.Center,
                ToolTip = "Поиск семейства по названию и его параметрам"
            };

            _tbSearch.GotFocus += (s, e) =>
            {
                if (_isWatermarkActive)
                {
                    _isApplyingWatermark = true;
                    _tbSearch.Text = "";
                    _tbSearch.Foreground = System.Windows.Media.Brushes.Black;
                    _isWatermarkActive = false;
                    _isApplyingWatermark = false;
                }
            };

            _tbSearch.LostFocus += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(_tbSearch.Text))
                    ActivateSearchWatermark();
            };

            _tbSearch.TextChanged += OnSearchTextChanged;
            _tbSearch.Loaded += (s, e) => ActivateSearchWatermark();
            ActivateSearchWatermark();

            _searchDebounceTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(200) };
            _searchDebounceTimer.Tick += (s, e) =>
            {
                _searchDebounceTimer.Stop();
                RefreshScenario();
            };

            Grid.SetRow(_tbSearch, 0);
            root.Children.Add(_tbSearch);

            bool isBimUi = IsBimAdmin(subDepartamentName);
            _cbStage = null;
            _cbProject = null;
            if (!isBimUi)
            {
                EnsureLookupsLoaded();

                var filtersPanel = new StackPanel
                {
                    Orientation = Orientation.Vertical,
                    Margin = new Thickness(0, 0, 0, 8),
                    HorizontalAlignment = HorizontalAlignment.Stretch,
                    VerticalAlignment = VerticalAlignment.Top
                };

                if (string.Equals(subDepartamentName?.Trim(), "АР", StringComparison.OrdinalIgnoreCase))
                {
                    _cbStage = new ComboBox
                    {
                        Height = 22,
                        Margin = new Thickness(0, 0, 0, 8),
                        HorizontalAlignment = HorizontalAlignment.Stretch,
                        VerticalAlignment = VerticalAlignment.Top,
                        IsEditable = false 
                    };
                    BindStageCombo_DefaultId1(_cbStage);
                    _cbStage.SelectionChanged += (s, e) =>
                    {
                        ClearCurrentDeptSelection(); 
                        RefreshScenario();
                        UpdateSelectionButtonsState();
                    };

                    filtersPanel.Children.Add(_cbStage);
                }

                _cbProject = new ComboBox
                {
                    Height = 22,
                    Margin = new Thickness(0, 0, 0, 0),
                    HorizontalAlignment = HorizontalAlignment.Stretch,
                    VerticalAlignment = VerticalAlignment.Top,
                    IsEditable = false
                };
                BindProjectCombo_DefaultId1(_cbProject);
                _cbProject.SelectionChanged += (s, e) =>
                {
                    ClearCurrentDeptSelection();
                    RefreshScenario();
                    UpdateSelectionButtonsState();
                };

                filtersPanel.Children.Add(_cbProject);

                Grid.SetRow(filtersPanel, 1);
                root.Children.Add(filtersPanel);
            }

            FrameworkElement scenarioUI = BuildScenarioUI(subDepartamentName, _records);
            Grid.SetRow(scenarioUI, 2);
            root.Children.Add(scenarioUI);

            if (!IsBimAdmin(subDepartamentName))
            {
                root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                var actionsPanel = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Left,
                    Margin = new Thickness(0, 10, 0, 0)
                };

                _btnOpenInRevit = new Button
                {
                    Content = "Редактировать в Revit",
                    MinWidth = 145,
                    Height = 22,
                    Margin = new Thickness(0, 0, 8, 0),
                    Background = (System.Windows.Media.Brush)new BrushConverter().ConvertFromString("#90EE90"),
                    IsEnabled = false
                };
                _btnOpenInRevit.Click += OnOpenInRevitClick;

                _btnLoadIntoProject = new Button
                {
                    Content = "Загрузить в проект",
                    MinWidth = 130,
                    Height = 22,
                    Background = (System.Windows.Media.Brush)new BrushConverter().ConvertFromString("#90EE90"),
                    IsEnabled = false
                };
                _btnLoadIntoProject.Click += OnLoadIntoProjectClick;

                actionsPanel.Children.Add(_btnOpenInRevit);
                actionsPanel.Children.Add(_btnLoadIntoProject);

                Grid.SetRow(actionsPanel, 3);
                root.Children.Add(actionsPanel);
            }

            _scenarioContent = scenarioUI;
            MainArea.Child = root;
        }

        // Обработчик собыйтий. Фильтр семейств по названию
        private void OnSearchTextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isApplyingWatermark || _isWatermarkActive) return;
            if (_searchDebounceTimer != null)
            {
                _searchDebounceTimer.Stop();
                _searchDebounceTimer.Start();
            }
            else
            {
                RefreshScenario();
            }
        }

        // Вспомогательный метод поиска. Распределение по отделам
        private void RefreshScenario()
        {
            if (_bimRootPanel != null)
            {
                RebuildBimContent();
                return;
            }

            if (_universalRootPanel != null)
            {
                RebuildUniversalContent();
                return;
            }
        }

        // Вспомогательный метод поиска. Активация подсказки в поле поиска
        private void ActivateSearchWatermark()
        {
            _isApplyingWatermark = true;
            _isWatermarkActive = true;
            _tbSearch.Foreground = System.Windows.Media.Brushes.Gray;
            _tbSearch.Text = SEARCH_WATERMARK;
            _isApplyingWatermark = false;
        }

        // UI в зависимости от подразделения
        private FrameworkElement BuildScenarioUI(string dep, List<FamilyManagerRecord> records)
        {
            if (IsBimAdmin(dep))
            {
                _universalRootPanel = null;
                _universalDepartment = null;
                _cbStage = null;
                _cbProject = null;

                var scroll = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
                _bimScrollViewer = scroll;
                _bimRootPanel = new StackPanel { Margin = new Thickness(0) };
                scroll.Content = _bimRootPanel;

                RebuildBimContent();
                return scroll;
            }
            else
            {
                _bimRootPanel = null;
                _bimScrollViewer = null;

                var scroll = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
                _univScrollViewer = scroll;
                var panel = new StackPanel { Margin = new Thickness(0) };
                scroll.Content = panel;

                _universalRootPanel = panel;  
                _universalDepartment = dep;    

                RebuildUniversalContent();  
                return scroll;
            }
        }

        // Вспомогательный метод интерфейса. Формирование имени
        private static string SafeFileName(string fullPath)
        {
            if (string.IsNullOrWhiteSpace(fullPath)) return "(нет имени)";
            try
            {
                return System.IO.Path.GetFileName(fullPath);
            }
            catch { return fullPath; }
        }

        // Интерфейс для BIM-отдела. Общее
        private void RebuildBimContent()
        {
            if (_bimRootPanel == null) return;

            CaptureBimUiState();

            _bimRootPanel.Children.Clear();

            var all = _records ?? new List<FamilyManagerRecord>();

            string q = _isWatermarkActive ? null : _tbSearch?.Text?.Trim();
            if (!string.IsNullOrEmpty(q))
            {
                all = all.Where(r => MatchesSearchByIndex(r, q)).ToList();
            }

            Func<string, string> norm = s => (s ?? "").Trim().ToUpperInvariant();
            var catNew = all.Where(r => norm(r.Status) == "NEW").ToList();
            var catAbsentError = all.Where(r =>
            {
                var st = norm(r.Status);
                return st == "ABSENT" || st == "ERROR";
            }).ToList();

            var catNotDepartament = all.Where(r => r.Departament == null && norm(r.Status) != "ABSENT" && norm(r.Status) != "ERROR" && norm(r.Status) != "IGNORE").ToList();
            var catOkWithBadMeta = all.Where(r => (norm(r.Status) != "ABSENT" && norm(r.Status) != "ERROR" && norm(r.Status) != "IGNORE") && r.Category == 1).ToList();
            var catOkMissingImportOrImage = all.Where(r => (norm(r.Status) != "ABSENT" && norm(r.Status) != "ERROR" && norm(r.Status) != "IGNORE") && (r.ImportInfo == null)).ToList();
            var catIgnored = all.Where(r => norm(r.Status) == "IGNORE").ToList();
            var catOkProcessed = all.Where(r => (norm(r.Status) == "OK") && r.Departament != null && r.Category != 1).ToList();

            _bimRootPanel.Children.Add(CreateCategoryExpander("НОВЫЕ СЕМЕЙСТВА", "bim_new.png", catNew, "Семейства, которые ранее были добавлены на диск и не были обработаны.\nДанные семейства не отображаются в списке у пользователей до изменения статуса на «OK» у семейства BIM-координатором.", key: "new"));
            _bimRootPanel.Children.Add(CreateCategoryExpander("ОШИБКИ / НЕ НАЙДЕН", "bim_error.png", catAbsentError, "Семейства, которые были удалены, или семейства, в которые параметры были экспортированы с ошибками.\nДанные семейства не отображаются в списке у пользователей до исправления ошибок BIM-координатором.", key: "errorabsent"));
            _bimRootPanel.Children.Add(CreateCategoryExpander("НЕТ ОТДЕЛА", "bim_error.png", catNotDepartament, "Семейства, которые не содержут информацию об отделе.\nДанные семейства не отображаются в списке у пользователей до указания отдела BIM-координатором.", key: "nodept"));
            _bimRootPanel.Children.Add(CreateCategoryExpander("НЕ УКАЗАНА КАТЕГОРИЯ", "bim_caution.png", catOkWithBadMeta, "Семейства, в которых не указан параметр КАТЕГОРИЯ.\nБез указания данного параметра семейство группируется в директории по-умолчанию.", key: "badmeta"));
            _bimRootPanel.Children.Add(CreateCategoryExpander("НЕТ СВОЙСТВ СЕМЕЙСТВА", "bim_caution.png", catOkMissingImportOrImage, "Семейства, в которых не указаны экспортируемые свойства семейства.\nБез указания данных параметров семейство не содержит описания о себе (из файла).", key: "missingimport"));
            _bimRootPanel.Children.Add(CreateCategoryExpander("АРХИВ", "bim_ignore.png", catIgnored, "Семейства, помеченные к игнорированию. Не отображаются у пользователей.", key: "ignored"));
            _bimRootPanel.Children.Add(CreateCategoryExpander("ОБРАБОТАННЫЕ СЕМЕЙСТВА", "bim_ok.png", catOkProcessed, "Полностью обработанные семейства со статусом «OK». Отдел указан, базовые поля заполнены, свойства из файла присутствуют.", key: "ok"));

            _bimRootPanel.UpdateLayout();
            RestoreBimUiStateAfterLayout();
        }

        // Интерфейс для BIM-отдела. Раскрывающийся список
        private Expander CreateCategoryExpander(string title, string iconFileName, List<FamilyManagerRecord> items, string tooltip, string key)
        {
            var panel = new StackPanel();
            foreach (var r in items.OrderBy(i => SafeFileName(i.FullPath), StringComparer.OrdinalIgnoreCase))
                panel.Children.Add(CreateRow(r));

            var header = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };

            var asm = this.GetType().Assembly;
            var resName = $"{asm.GetName().Name}.Imagens.FamilyManager.{iconFileName}";

            using (var s = asm.GetManifestResourceStream(resName))
            {
                if (s == null)
                {
                    var have = string.Join(", ", asm.GetManifestResourceNames().Where(n => n.Contains("Imagens")));
                    throw new IOException($"Не найден Embedded Resource '{resName}'. Есть: {have}");
                }

                var bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.StreamSource = s;
                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.EndInit();
                bmp.Freeze();

                header.Children.Add(new System.Windows.Controls.Image { Width = 16, Height = 16, Margin = new Thickness(3, 0, 6, 0), Source = bmp });
            }

            var tb = new TextBlock { VerticalAlignment = VerticalAlignment.Center, TextTrimming = TextTrimming.CharacterEllipsis };
            tb.Inlines.Add(new Run(title) { FontWeight = FontWeights.Bold });
            tb.Inlines.Add(new Run($" ({items?.Count ?? 0})"));
            header.Children.Add(tb);

            ToolTipService.SetShowDuration(header, 30000);
            ToolTipService.SetInitialShowDelay(header, 300);
            ToolTipService.SetBetweenShowDelay(header, 150);
            ToolTipService.SetToolTip(header, new ToolTip { Content = tooltip });

            bool isExpanded = _bimExpandState.TryGetValue(key, out var saved) ? saved : false;

            var exp = new Expander
            {
                Header = header,
                IsExpanded = isExpanded,
                Margin = new Thickness(0, 0, 0, 6),
                Content = panel,
                Tag = key
            };

            exp.Expanded += (s, e) => _bimExpandState[key] = true;
            exp.Collapsed += (s, e) => _bimExpandState[key] = false;

            return exp;
        }

        // Интерфейс для BIM-отдела. Содержимое раскрывающегося списка
        private FrameworkElement CreateRow(FamilyManagerRecord rec)
        {
            var grid = new Grid { Margin = new Thickness(16, 8, 0, 2) };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var name = SafeFileName(rec.FullPath);

            var btn = new Button
            {
                Padding = new Thickness(0),
                Background = System.Windows.Media.Brushes.Transparent,
                BorderThickness = new Thickness(0),
                HorizontalContentAlignment = HorizontalAlignment.Left,

                ToolTip = rec.FullPath,
                Tag = rec.ID
            };

            btn.Content = new TextBlock
            {
                Text = name,
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis
            };

            btn.Click += OnRowMainClick;

            Grid.SetColumn(btn, 0);
            grid.Children.Add(btn);

            return grid;
        }

        // Интерфейс для BIM-отдела. Редактировать информацию о семействе
        private void OnRowMainClick(object sender, RoutedEventArgs e)
        {
            var idObj = (sender as System.Windows.Controls.Button)?.Tag;
            var idText = idObj?.ToString() ?? "null";
            OpenFamilyEditorLoop(idText);
        }

        // Логика открытия окна. BIM
        private void OpenFamilyEditorLoop(string idText)
        {
            var owner = Window.GetWindow(this) ?? Application.Current?.MainWindow;

            while (true)
            {
                var win = new KPLN_Tools.Forms.FamilyManagerEditBIM(idText)
                {
                    Owner = owner,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    ShowInTaskbar = false,
                    Topmost = true
                };

                if (win.ShowDialog() != true)
                    break;

                string openFormat = win.OpenFormat;
                string filePathFI = win.filePathFI;

                if (openFormat == "OpenFamily" || openFormat == "OpenFamilyInProject")
                {
                    if (_uiapp == null)
                    {
                        TaskDialog.Show("Ошибка", "Не удалось установить UIApplication.");
                        break;
                    }

                    string path = filePathFI;
                    if (!System.IO.File.Exists(path))
                    {
                        TaskDialog.Show("Ошибка", $"Файл не найден: {path}");
                        break;
                    }

                    try
                    {
                        if (openFormat == "OpenFamily")
                        {
                            _uiapp.OpenAndActivateDocument(path);
                        }
                        else
                        {
                            ExternalEventsHost.LoadFamilyHandler.FilePath = path;
                            ExternalEventsHost.LoadFamilyEvent.Raise();
                        }
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show("Ошибка", ex.Message);
                    }

                    break;
                }
                else if (openFormat == "UpdateFamily")
                {
                    bool ok = TryUpdateImportInfoForRecord(idText);
                    if (ok)
                        ReloadFromDbAndRefreshUI();

                    continue;
                }
                else
                {
                    if (win.DeleteStatus)
                    {
                        if (int.TryParse(idText, out int idToDelete))
                        {
                            try { DeleteRecordFromDatabase(idToDelete); }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Не удалось удалить запись: " + ex.Message,
                                    "Family Manager", MessageBoxButton.OK, MessageBoxImage.Error);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Не выбран элемент для удаления.",
                                "Family Manager", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                    }
                    else
                    {
                        var rec = win.ResultRecord;
                        if (rec == null)
                        {
                            MessageBox.Show("Нет данных для сохранения.",
                                "Family Manager", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                        else
                        {
                            try { FamilyManagerEditBIM.SaveRecordToDatabase(rec); }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Не удалось сохранить запись: " + ex.Message,
                                    "Family Manager", MessageBoxButton.OK, MessageBoxImage.Error);
                            }
                        }
                    }

                    ReloadFromDbAndRefreshUI();
                    break;
                }
            }
        }



        // Поиск. Только по значениям (и имени файла), из _searchIndex
        private bool MatchesSearchByIndex(FamilyManagerRecord r, string query)
        {
            if (string.IsNullOrWhiteSpace(query)) return true;
            if (r == null) return false;

            var q = query.ToUpperInvariant();

            if (_searchIndex != null && _searchIndex.TryGetValue(r.ID, out var text) && !string.IsNullOrEmpty(text))
                return text.Contains(q);

            return false;
        }



        private void OpenFamilyUserLoop(string idText)
        {
            var owner = Window.GetWindow(this) ?? Application.Current?.MainWindow;

            while (true)
            {
                var win = new KPLN_Tools.Forms.FamilyManagerEditUser(idText)
                {
                    Owner = owner,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    ShowInTaskbar = false,
                    Topmost = true
                };

                if (win.ShowDialog() != true)
                    break;

                string openFormat = win.OpenFormat;
                string filePathFI = win.filePathFI;

                if (openFormat == "OpenFamily" || openFormat == "OpenFamilyInProject")
                {
                    if (_uiapp == null)
                    {
                        TaskDialog.Show("Ошибка", "Не удалось установить UIApplication.");
                        break;
                    }

                    string path = filePathFI;
                    if (!System.IO.File.Exists(path))
                    {
                        TaskDialog.Show("Ошибка", $"Файл не найден: {path}");
                        break;
                    }

                    try
                    {
                        if (openFormat == "OpenFamily")
                        {
                            _uiapp.OpenAndActivateDocument(path);
                        }
                        else
                        {
                            ExternalEventsHost.LoadFamilyHandler.FilePath = path;
                            ExternalEventsHost.LoadFamilyEvent.Raise();
                        }
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show("Ошибка", ex.Message);
                    }

                    break;
                }            
            }
        }

        /// <summary>
        /// Обновляет IMPORT_INFO для записи FamilyManager по текстовому ID.
        /// Делает все проверки статуса/отдела, открывает RFA, читает параметры, пишет JSON в БД.
        /// Возвращает true, если IMPORT_INFO был обновлён.
        /// </summary>
        private bool TryUpdateImportInfoForRecord(string idText)
        {
            // ID
            if (!int.TryParse(idText, out int famId))
            {
                MessageBox.Show("Неверный ID записи в БД.", "Family Manager", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            // FamilyManager (Status, FullPath, Departament)
            var recRow = GetFamilyRowById(DB_PATH, famId);
            if (recRow == null)
            {
                MessageBox.Show($"Запись {famId} не найдена в БД.", "Family Manager", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            // DEPARTAMENT
            string dep = (recRow.Departament ?? "").Trim();
            if (string.IsNullOrWhiteSpace(dep))
            {
                MessageBox.Show("Невозможно обновлять данные, не указав отдел.\nОперация пропущена.", "Family Manager", MessageBoxButton.OK, MessageBoxImage.Information);
                return false;
            }

            // Проверка пути
            string famPath = recRow.FullPath;
            if (string.IsNullOrWhiteSpace(famPath) || !System.IO.File.Exists(famPath))
            {
                MessageBox.Show($"Файл семейства не найден:\n{famPath}", "Family Manager", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            // Чтение параметров семейства и сборка JSON
            string json;
            try
            {
                if (_uiapp == null)
                {
                    TaskDialog.Show("Ошибка", "Не удалось установить UIApplication.");
                    return false;
                }
                json = ReadImportInfoFromFamily(_uiapp, famPath, dep);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка чтения параметров из семейства:\n" + ex.Message, "Family Manager", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
            if (string.IsNullOrWhiteSpace(json))
            {
                MessageBox.Show("Ошибка обработки параметров.", "Family Manager", MessageBoxButton.OK, MessageBoxImage.Information);
                return false;
            }

            // Запись JSON в БД
            try
            {
                int rows = UpdateImportInfoJson(DB_PATH, famId, json);
                if (rows > 0)
                {
                    return true;
                }
                else
                {
                    MessageBox.Show("Запись в БД не изменёна (возможно, запись не найдена).", "Family Manager", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка записи JSON в БД:\n" + ex.Message, "Family Manager", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        // IMPORT_INFO. Локальная модель строки из БД для поиска по ID
        internal class FamilyRowLite
        {
            public int ID;
            public string Status;
            public string FullPath;
            public string Departament;
        }

        // IMPORT_INFO. Чтение одной записи по ID из БД
        internal static FamilyRowLite GetFamilyRowById(string dbPath, int id)
        {
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand(
                    "SELECT ID, STATUS, FULLPATH, DEPARTAMENT FROM FamilyManager WHERE ID = @id LIMIT 1;", conn))
                {
                    cmd.Parameters.AddWithValue("@id", id);
                    using (var rd = cmd.ExecuteReader())
                    {
                        if (!rd.Read()) return null;
                        return new FamilyRowLite
                        {
                            ID = Convert.ToInt32(rd["ID"]),
                            Status = rd["STATUS"] as string,
                            FullPath = rd["FULLPATH"] as string,
                            Departament = rd["DEPARTAMENT"] as string
                        };
                    }
                }
            }
        }

        // IMPORT_INFO. Открывает семейство в зависимости от отдела и возвращает JSON
        internal static string ReadImportInfoFromFamily(UIApplication uiapp, string familyPath, string department, string rteTemplateIfNeeded = null)
        {
            var kind = DetectDepartmentKind(department);
            if (kind == DepartmentKind.Unknown)
                return "";

            Document famDoc = null;
            try
            {
                famDoc = uiapp.Application.OpenDocumentFile(familyPath);
                if (famDoc == null || !famDoc.IsFamilyDocument)
                    throw new InvalidOperationException("Файл не является документом семейства.");

                var fm = famDoc.FamilyManager;
                if (fm == null || fm.Types == null || fm.Types.Size == 0)
                    return "";

                switch (kind)
                {
                    case DepartmentKind.AR: return ExtractJson_AR(famDoc, fm);
                    case DepartmentKind.KR: return ExtractJson_KR(famDoc, fm);
                    case DepartmentKind.OViK: return ExtractJson_OViK(famDoc, fm);
                    case DepartmentKind.VK: return ExtractJson_VK(famDoc, fm);
                    case DepartmentKind.EOM: return ExtractJson_EOM(famDoc, fm);
                    case DepartmentKind.SS: return ExtractJson_SS(famDoc, fm);
                    default: return "";
                }
            }
            finally
            {
                if (famDoc != null) { try { famDoc.Close(false); } catch { } }
            }
        }

        // IMPORT_INFO. Список отделов
        private enum DepartmentKind { Unknown, AR, KR, OViK, VK, EOM, SS }

        // IMPORT_INFO. Отдел - Параметры
        private static DepartmentKind DetectDepartmentKind(string dep)
        {
            if (string.IsNullOrWhiteSpace(dep)) return DepartmentKind.Unknown;
            if (ContainsInvariant(dep, "АР")) return DepartmentKind.AR;
            if (ContainsInvariant(dep, "КР")) return DepartmentKind.KR;
            if (ContainsInvariant(dep, "ОВиК")) return DepartmentKind.OViK;
            if (ContainsInvariant(dep, "ВК")) return DepartmentKind.VK;
            if (ContainsInvariant(dep, "ЭОМ")) return DepartmentKind.EOM;
            if (ContainsInvariant(dep, "СС")) return DepartmentKind.SS;
            return DepartmentKind.Unknown;
        }

        // IMPORT_INFO. Отдел - Параметры. Проверка вхождения
        private static bool ContainsInvariant(string haystack, string needle)
        {
            return haystack?.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /////////////////// IMPORT_INFO. Список всех отделов
        ////////////////////// 
        // --- АР ---
        private static string ExtractJson_AR(Document doc, Autodesk.Revit.DB.FamilyManager fm)
        {
            var pNote = fm.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            string note = JoinAllValuesByTypes(fm, pNote);
            return "{"
                 + "\"Примечание\":\"" + JsonEscape(note) + "\""
                 + "}";
        }
        // --- КР ---
        private static string ExtractJson_KR(Document doc, Autodesk.Revit.DB.FamilyManager fm)
        {
            var pNote = fm.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            string note = JoinAllValuesByTypes(fm, pNote);
            return "{"
                 + "\"Примечание\":\"" + JsonEscape(note) + "\""
                 + "}";
        }
        // --- ОВиК ---
        private static string ExtractJson_OViK(Document doc, Autodesk.Revit.DB.FamilyManager fm)
        {
            var pNote = fm.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            var pDescr = GetFamilyParamByNameThenGuid(fm, "КП_О_Наименование", new Guid("f194bf60-b880-4217-b793-1e0c30dda5e9"));
            var pMark = GetFamilyParamByNameThenGuid(fm, "КП_О_Марка", new Guid("2204049c-d557-4dfc-8d70-13f19715e46d"));
            var pManuf = GetFamilyParamByNameThenGuid(fm, "КП_О_Завод-изготовитель", new Guid("a8cdbf7b-d60a-485e-a520-447d2055f351"));

            string note = JoinAllValuesByTypes(fm, pNote);
            string descr = JoinAllValuesByTypes(fm, pDescr);
            string mark = JoinAllValuesByTypes(fm, pMark);
            string manuf = JoinAllValuesByTypes(fm, pManuf);

            return "{"
                 + "\"Описание\":\"" + JsonEscape(descr) + "\","
                 + "\"Марка\":\"" + JsonEscape(mark) + "\","
                 + "\"Производитель\":\"" + JsonEscape(manuf) + "\","
                 + "\"Примечание\":\"" + JsonEscape(note) + "\""
                 + "}";
        }
        // --- ВК ---
        private static string ExtractJson_VK(Document doc, Autodesk.Revit.DB.FamilyManager fm)
        {
            var pNote = fm.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            var pDescr = GetFamilyParamByNameThenGuid(fm, "КП_О_Наименование", new Guid("f194bf60-b880-4217-b793-1e0c30dda5e9"));
            var pMark = GetFamilyParamByNameThenGuid(fm, "КП_О_Марка", new Guid("2204049c-d557-4dfc-8d70-13f19715e46d"));
            var pManuf = GetFamilyParamByNameThenGuid(fm, "КП_О_Завод-изготовитель", new Guid("a8cdbf7b-d60a-485e-a520-447d2055f351"));

            string note = JoinAllValuesByTypes(fm, pNote);
            string descr = JoinAllValuesByTypes(fm, pDescr);
            string mark = JoinAllValuesByTypes(fm, pMark);
            string manuf = JoinAllValuesByTypes(fm, pManuf);

            return "{"
                 + "\"Описание\":\"" + JsonEscape(descr) + "\","
                 + "\"Марка\":\"" + JsonEscape(mark) + "\","
                 + "\"Производитель\":\"" + JsonEscape(manuf) + "\","
                 + "\"Примечание\":\"" + JsonEscape(note) + "\""
                 + "}";
        }
        // --- ЭОМ ---
        private static string ExtractJson_EOM(Document doc, Autodesk.Revit.DB.FamilyManager fm)
        {
            var pNote = fm.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            var pDescr = GetFamilyParamByNameThenGuid(fm, "КП_О_Наименование", new Guid("f194bf60-b880-4217-b793-1e0c30dda5e9"));
            var pMark = GetFamilyParamByNameThenGuid(fm, "КП_О_Марка", new Guid("2204049c-d557-4dfc-8d70-13f19715e46d"));
            var pManuf = GetFamilyParamByNameThenGuid(fm, "КП_О_Завод-изготовитель", new Guid("a8cdbf7b-d60a-485e-a520-447d2055f351"));

            string note = JoinAllValuesByTypes(fm, pNote);
            string descr = JoinAllValuesByTypes(fm, pDescr);
            string mark = JoinAllValuesByTypes(fm, pMark);
            string manuf = JoinAllValuesByTypes(fm, pManuf);

            return "{"
                 + "\"Описание\":\"" + JsonEscape(descr) + "\","
                 + "\"Марка\":\"" + JsonEscape(mark) + "\","
                 + "\"Производитель\":\"" + JsonEscape(manuf) + "\","
                 + "\"Примечание\":\"" + JsonEscape(note) + "\""
                 + "}";
        }

        // --- СС ---
        private static string ExtractJson_SS(Document doc, Autodesk.Revit.DB.FamilyManager fm)
        {
            var pNote = fm.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            var pDescr = GetFamilyParamByNameThenGuid(fm, "КП_О_Наименование", new Guid("f194bf60-b880-4217-b793-1e0c30dda5e9"));
            var pMark = GetFamilyParamByNameThenGuid(fm, "КП_О_Марка", new Guid("2204049c-d557-4dfc-8d70-13f19715e46d"));
            var pManuf = GetFamilyParamByNameThenGuid(fm, "КП_О_Завод-изготовитель", new Guid("a8cdbf7b-d60a-485e-a520-447d2055f351"));

            string note = JoinAllValuesByTypes(fm, pNote);
            string descr = JoinAllValuesByTypes(fm, pDescr);
            string mark = JoinAllValuesByTypes(fm, pMark);
            string manuf = JoinAllValuesByTypes(fm, pManuf);

            return "{"
                 + "\"Описание\":\"" + JsonEscape(descr) + "\","
                 + "\"Марка\":\"" + JsonEscape(mark) + "\","
                 + "\"Производитель\":\"" + JsonEscape(manuf) + "\","
                 + "\"Примечание\":\"" + JsonEscape(note) + "\""
                 + "}";
        }

        // IMPORT_INFO. Поиск FamilyParameter по Ищем по имени, иначе по GUID
        private static FamilyParameter GetFamilyParamByNameThenGuid(Autodesk.Revit.DB.FamilyManager fm, string exactName, Guid guid)
        {
            if (fm == null) return null;

            foreach (FamilyParameter fp in fm.Parameters)
            {
                var name = fp.Definition?.Name;
                if (!string.IsNullOrEmpty(name) && string.Equals(name, exactName, StringComparison.Ordinal))
                    return fp;
            }

            foreach (FamilyParameter fp in fm.Parameters)
            {
                try
                {
                    Guid g = fp.GUID;
                    if (g == guid)
                        return fp;
                }
                catch{}
            }

            return null;
        }

        // IMPORT_INFO. Все непустые уникальные значения параметра по всем типам, через запятую
        private static string JoinAllValuesByTypes(Autodesk.Revit.DB.FamilyManager fm, FamilyParameter param)
        {
            if (fm == null) return "";
            if (param == null) return "";
            if (fm.Types == null || fm.Types.Size == 0) return "";

            var seen = new HashSet<string>(StringComparer.Ordinal);
            var result = new List<string>();

            foreach (FamilyType t in fm.Types)
            {
                if (t == null) continue;
                var val = GetFamilyParamStringValue(t, param);
                if (string.IsNullOrWhiteSpace(val)) continue;

                if (seen.Add(val))
                    result.Add(val);
            }

            return string.Join(", ", result);
        }

        // IMPORT_INFO. Чтение строкового представления значения параметра у конкретного FamilyType
        private static string GetFamilyParamStringValue(FamilyType type, FamilyParameter fp)
        {
            if (type == null || fp == null) return null;

            try
            {
                switch (fp.StorageType)
                {
                    case StorageType.String:
                        return type.AsString(fp);

                    case StorageType.Double:
                        {
                            double? d = type.AsDouble(fp);
                            return d.HasValue
                                ? d.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)
                                : null;
                        }

                    case StorageType.Integer:
                        {
                            int? i = type.AsInteger(fp);
                            return i.HasValue
                                ? i.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)
                                : null;
                        }

                    case StorageType.ElementId:
                        {
                            ElementId eid = type.AsElementId(fp);
                            if (eid == null || eid == ElementId.InvalidElementId)
                                return null;
                            return eid.IntegerValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
                        }

                    default:
                        return null;
                }
            }
            catch
            {
                return null;
            }
        }

        // IMPORT_INFO. Экранирование для JSON
        private static string JsonEscape(string s)
        {
            if (s == null) return "";
            return s
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("\r", "\\r")
                .Replace("\n", "\\n")
                .Replace("\t", "\\t");
        }

        // IMPORT_INFO. Обновление IMPORT_INFO в БД
        internal static int UpdateImportInfoJson(string dbPath, int id, string json)
        {
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand(
                    "UPDATE FamilyManager SET IMPORT_INFO = @json WHERE ID = @id;", conn))
                {
                    cmd.Parameters.AddWithValue("@json", json);
                    cmd.Parameters.AddWithValue("@id", id);
                    return cmd.ExecuteNonQuery();
                }
            }
        }

        // Удаление записи из БД
        public static void DeleteRecordFromDatabase(int id)
        {
            if (id <= 0) throw new ArgumentOutOfRangeException(nameof(id));

            var cs = $"Data Source={DB_PATH};Version=3;Read Only=False;Foreign Keys=True;";

            using (var conn = new SQLiteConnection(cs))
            {
                conn.Open();
                using (var tr = conn.BeginTransaction())
                using (var cmd = conn.CreateCommand())
                {
                    cmd.Transaction = tr;
                    cmd.CommandText = "DELETE FROM FamilyManager WHERE ID = @id;";
                    cmd.Parameters.AddWithValue("@id", id);

                    var affected = cmd.ExecuteNonQuery();
                    tr.Commit();
                }
            }
        }

        // Интерфейс для BIM-отдела. Сохраняем стейт
        private void CaptureBimUiState()
        {
            if (_bimRootPanel == null) return;

            foreach (var exp in _bimRootPanel.Children.OfType<Expander>())
            {
                var key = exp.Tag as string;
                if (!string.IsNullOrEmpty(key))
                    _bimExpandState[key] = exp.IsExpanded;
            }

            if (_bimScrollViewer != null)
                _bimScrollOffset = _bimScrollViewer.VerticalOffset;
        }

        // Интерфейс для BIM-отдела. Сбрасываем стейт
        private void RestoreBimUiStateAfterLayout()
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (_bimScrollViewer != null)
                    _bimScrollViewer.ScrollToVerticalOffset(_bimScrollOffset);
            }), System.Windows.Threading.DispatcherPriority.Loaded);
        }

        // Интерфейс для BIM-отдела. Обновляем стейт
        private void ReloadFromDbAndRefreshUI()
        {
            try
            {
                var depUi = GetCurrentDepartment(); // <— только из UI
                if (string.Equals(depUi, BIM_ADMIN, StringComparison.OrdinalIgnoreCase))
                {
                    _records = LoadFamilyManagerRecords(DB_PATH);
                }
                else
                {
                    _records = LoadFamilyManagerRecordsForDepartment(DB_PATH, depUi);
                }

                RebuildSearchIndex();
                BuildMainArea(depUi); 
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка чтения БД", $"{ex}");
            }
        }

        // XAML. Настройки
        private void BtnSettings_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            var dlg = new SettingsDialog
            {
                WindowStartupLocation = WindowStartupLocation.CenterScreen
            };

            if (dlg.ShowDialog() == true)
            {
                if (dlg.Result == 1)
                {
                    if (!Directory.Exists(RFA_ROOT))
                    {
                        TaskDialog.Show("Ошибка", $"Папка не найдена:\n{RFA_ROOT}");
                        return;
                    }
                    if (!File.Exists(DB_PATH))
                    {
                        TaskDialog.Show("Ошибка", $"Файл БД не найден:\n{DB_PATH}");
                        return;
                    }

                    try
                    {
                        MainArea.Child = null;

                        var rfaFiles = EnumerateRfaFilesSkippingArchive(RFA_ROOT).ToList();
                        if (rfaFiles.Count == 0)
                        {
                            TaskDialog.Show("Индексация RFA", "Файлы *.rfa не найдены или все находятся в папках «Архив» :)");
                            return;
                        }

                        var (inserted, updated, deleted) = InsertRfaRecords(DB_PATH, rfaFiles);
                        TaskDialog.Show("Индексация RFA",
                            $"Добавлено: {inserted}\nУдалено: {deleted}");
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show("Ошибка", ex.Message);
                    }
                }
                else if (dlg.Result == 2)
                {
                    var optDlg = new FamilyManagerSettingsImport
                    {
                        WindowStartupLocation = WindowStartupLocation.CenterScreen
                    };

                    if (optDlg.ShowDialog() == true)
                    {
                        bool doDepartment = optDlg.DoDepartment;
                        bool doImport = optDlg.DoImportParams;
                        bool doImage = optDlg.DoFamilyImage;

                        if (!Directory.Exists(RFA_ROOT))
                        {
                            TaskDialog.Show("Ошибка", $"Папка не найдена:\n{RFA_ROOT}");
                            return;
                        }
                        if (!File.Exists(DB_PATH))
                        {
                            TaskDialog.Show("Ошибка", $"Файл БД не найден:\n{DB_PATH}");
                            return;
                        }

                        try
                        {
                            MainArea.Child = null;

                            string dubugInfo = "";
                            int updatedCount = 0;

                            var sb = new System.Text.StringBuilder();

                            if (doDepartment)
                            {
                                updatedCount = UpdateDepartmentsByPath(DB_PATH);
                                dubugInfo += $"Обновлено значение «Отдел» в БД: {updatedCount}\n";
                            }
                            if (doImage)
                            {
                                int updatedImages = UpdateImagesByPath(DB_PATH);
                                dubugInfo += $"Обновлено значение «Изображение» в БД: {updatedImages}\n";
                            }

                            if (doImport)
                            {
                                ExternalEventsHost.EnsureCreated();

                                var h = ExternalEventsHost.BulkPagedUpdateHandler;
                                h.DbPath = DB_PATH;
                                h.PageSize = 200;

                                h.Completed = (handler) =>
                                {
                                    sb.AppendLine($"Обновлено «IMPORT_INFO»: {handler.Updated}.");
                                    sb.AppendLine($"Ошибок: {handler.Errors}.\n");

                                    if (handler.Errors > 0 && handler.ErrorLog != null && handler.ErrorLog.Count > 0)
                                    {
                                        try
                                        {
                                            var wantSave = System.Windows.MessageBox.Show(
                                                "Во время обновления возникли ошибки. Сохранить отчёт с перечнем FULLPATH и описанием ошибок?", "Family Manager. Отчет об ошибках",
                                                MessageBoxButton.YesNo, MessageBoxImage.Question);

                                            if (wantSave == MessageBoxResult.Yes)
                                            {
                                                var sfd = new Microsoft.Win32.SaveFileDialog
                                                {
                                                    Title = "Сохранить отчет об ошибках",
                                                    Filter = "Текстовый файл (*.txt)|*.txt",
                                                    FileName = $"FamilyManager_Errors_{DateTime.Now:yyyyMMdd_HHmm}.txt"
                                                };

                                                if (sfd.ShowDialog() == true)
                                                {
                                                    var lines = new List<string>();
                                                    lines.Add($"Дата: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                                                    lines.Add($"База: {DB_PATH}");
                                                    lines.Add($"Всего ошибок: {handler.Errors}");
                                                    lines.Add("");
                                                    lines.Add("ID\tINFO\tFULLPATH");
                                                    foreach (var er in handler.ErrorLog)
                                                    {
                                                        string id = er.Id.ToString();
                                                        string msg = er.Message?.Replace("\r", " ").Replace("\n", " ");
                                                        string full = string.IsNullOrWhiteSpace(er.FullPath) ? "(нет пути)" : er.FullPath;                                                        
                                                        lines.Add($"{id}\t{msg}\t{full}");
                                                    }

                                                    File.WriteAllLines(sfd.FileName, lines, Encoding.UTF8);
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            System.Windows.MessageBox.Show("Не удалось сохранить отчет об ошибках:\n" + ex.Message, "Family Manager", MessageBoxButton.OK, MessageBoxImage.Warning);
                                        }
                                    }

                                    try
                                    {
                                        TaskDialog.Show("Обновлены данные в БД", sb.ToString());
                                    }
                                    catch
                                    {
                                        System.Windows.MessageBox.Show(sb.ToString(), "Обновлены данные в БД", MessageBoxButton.OK);
                                    }

                                    ReloadFromDbAndRefreshUI();
                                };

                                h.IsRunning = true;
                                ExternalEventsHost.BulkPagedUpdateEvent.Raise();
                                return;
                            }
                            if (doDepartment || doImage)
                            {
                                dubugInfo += $"Опперация выполнена успешно.";
                                TaskDialog.Show("Обновлены данные в БД", $"{dubugInfo}");
                            }
                            else
                            {
                                TaskDialog.Show("Информация", $"Не выбрано ни одного параметра.");
                            }
                        }
                        catch (Exception ex)
                        {
                            TaskDialog.Show("Ошибка", ex.Message);
                        }
                    }
                }
                else if (dlg.Result == 3)
                {
                    var win = new FamilyManagerSettingDB
                    {
                        WindowStartupLocation = WindowStartupLocation.CenterScreen
                    };

                    win.ShowDialog();
                }
            }
        }

        /// --- Обновлеие ID, FILEPATH, STATUS в FamilyManager
        // Обход файлов .rfa 
        private static IEnumerable<string> EnumerateRfaFilesSkippingArchive(string root)
        {
            var stack = new Stack<string>();
            if (Directory.Exists(root)) stack.Push(root);

            while (stack.Count > 0)
            {
                string dir = stack.Pop();
                string dirName = Path.GetFileName(dir) ?? dir;

                if ((dirName.IndexOf("архив", StringComparison.OrdinalIgnoreCase) >= 0)
                    || (dirName.IndexOf("8_Библиотека семейств Самолета", StringComparison.OrdinalIgnoreCase) >= 0))
                    continue;

                IEnumerable<string> subdirs = Enumerable.Empty<string>();
                IEnumerable<string> files = Enumerable.Empty<string>();

                try { subdirs = Directory.EnumerateDirectories(dir); } catch { }
                try { files = Directory.EnumerateFiles(dir, "*.rfa", SearchOption.TopDirectoryOnly); } catch { }

                foreach (var f in files)
                {
                    if (ShouldSkipByBackupSuffix(f))
                        continue;

                    yield return f;
                }

                foreach (var sd in subdirs) stack.Push(sd);
            }
        }

        /// --- Обновлеие ID, FILEPATH, STATUS в FamilyManager
        // Игнорирование дуюликатов семейств
        private static bool ShouldSkipByBackupSuffix(string path)
        {
            // Ищем суффикс вида ".00dd" прямо перед расширением .rfa
            var nameNoExt = Path.GetFileNameWithoutExtension(path);
            if (string.IsNullOrEmpty(nameNoExt) || nameNoExt.Length < 5)
                return false;

            int i = nameNoExt.Length - 5;
            if (nameNoExt[i] != '.')
                return false;

            char c0 = nameNoExt[i + 1];
            char c1 = nameNoExt[i + 2];
            char c2 = nameNoExt[i + 3];
            char c3 = nameNoExt[i + 4];

            return c0 == '0' && c1 == '0'
                   && c2 >= '0' && c2 <= '9'
                   && c3 >= '0' && c3 <= '9';
        }

        /// --- Обновлеие ID, FILEPATH, STATUS в FamilyManager
        // Запись данных в БД (FamilyManager). ID, FILEPATH, STATUS
        private static (int inserted, int updated, int deleted) InsertRfaRecords(string dbPath, List<string> rfaPaths)
        {
            int inserted = 0;
            int updated = 0;
            int deleted = 0;

            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();

                using (var tx = conn.BeginTransaction())
                {
                    var byFullPath = new Dictionary<string, (long id, string lmDate, string status)>(StringComparer.OrdinalIgnoreCase);
                    using (var cmdSel = new SQLiteCommand("SELECT ID, FULLPATH, LM_DATE, STATUS FROM FamilyManager;", conn, tx))
                    using (var rd = cmdSel.ExecuteReader())
                    {
                        while (rd.Read())
                        {
                            long id = rd.GetInt64(0);
                            string fp = rd.GetString(1);
                            string lm = rd.IsDBNull(2) ? "" : rd.GetString(2);
                            string st = rd.IsDBNull(3) ? "" : rd.GetString(3);
                            byFullPath[fp] = (id, lm, st);
                        }
                    }

                    long nextId;
                    using (var cmdMax = new SQLiteCommand("SELECT COALESCE(MAX(ID), 0) FROM FamilyManager;", conn, tx))
                    {
                        var o = cmdMax.ExecuteScalar();
                        nextId = (o == null || o == DBNull.Value) ? 1L : Convert.ToInt64(o) + 1L;
                    }

                    using (var cmdUpdateStatusOnly = new SQLiteCommand(
                               @"UPDATE FamilyManager
                         SET STATUS = @ustatus
                         WHERE ID = @uid;", conn, tx))
                    using (var cmdInsert = new SQLiteCommand(
                               @"INSERT INTO FamilyManager (ID, STATUS, FULLPATH, LM_DATE)
                         VALUES (@id, @status, @fullpath, @lmdate);", conn, tx))
                    {
                        var pUStatus = cmdUpdateStatusOnly.Parameters.Add("@ustatus", System.Data.DbType.String);
                        var pUId = cmdUpdateStatusOnly.Parameters.Add("@uid", System.Data.DbType.Int64);

                        var pId = cmdInsert.Parameters.Add("@id", System.Data.DbType.Int64);
                        var pStatus = cmdInsert.Parameters.Add("@status", System.Data.DbType.String);
                        var pPath = cmdInsert.Parameters.Add("@fullpath", System.Data.DbType.String);
                        var pLmDate = cmdInsert.Parameters.Add("@lmdate", System.Data.DbType.String);

                        foreach (var kv in byFullPath.ToList())
                        {
                            string path = kv.Key;
                            var rec = kv.Value;

                            bool exists = false;
                            try { exists = File.Exists(path); } catch { }

                            if (!exists && !string.Equals(rec.status, "ABSENT", StringComparison.OrdinalIgnoreCase))
                            {
                                pUStatus.Value = "ABSENT";
                                pUId.Value = rec.id;
                                cmdUpdateStatusOnly.ExecuteNonQuery();
                                deleted++;

                                byFullPath[path] = (rec.id, rec.lmDate, "ABSENT");
                            }
                        }

                        foreach (var path in rfaPaths)
                        {
                            if (byFullPath.TryGetValue(path, out var existed))
                            {
                                if (string.Equals(existed.status, "IGNORE", StringComparison.OrdinalIgnoreCase))
                                    continue;

                                if (string.Equals(existed.status, "OK", StringComparison.OrdinalIgnoreCase))
                                    continue;

                                if (!string.Equals(existed.status, "NEW", StringComparison.OrdinalIgnoreCase))
                                {
                                    pUStatus.Value = "NEW";
                                    pUId.Value = existed.id;
                                    cmdUpdateStatusOnly.ExecuteNonQuery();
                                    updated++;

                                    byFullPath[path] = (existed.id, existed.lmDate, "NEW");
                                }
                            }
                            else
                            {
                                pId.Value = nextId;
                                pStatus.Value = "NEW";
                                pPath.Value = path;
                                pLmDate.Value = "";
                                cmdInsert.ExecuteNonQuery();
                                inserted++;

                                byFullPath[path] = (nextId, "", "NEW");
                                nextId++;
                            }
                        }
                    }

                    tx.Commit();
                }
            }

            return (inserted, updated, deleted);
        }

        /// --- Обновлеие DEPARTAMENT в FamilyManager
        private static int UpdateDepartmentsByPath(string dbPath)
        {
            string Like(string s)
            {
                if (string.IsNullOrEmpty(s)) return "%";
                return s.Replace("|", "||").Replace("_", "|_").Replace("%", "|%") + "%";
            }

            string ContainsLike(string s)
            {
                if (string.IsNullOrEmpty(s)) return "%";
                return "%" + s.Replace("|", "||").Replace("_", "|_").Replace("%", "|%") + "%";
            }

            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                using (var tx = conn.BeginTransaction())
                using (var cmd = conn.CreateCommand())
                {
                    cmd.Transaction = tx;

                    cmd.CommandText = @"
                UPDATE FamilyManager
                    SET 
                      DEPARTAMENT = CASE
                        WHEN (FULLPATH LIKE @bimTemplates ESCAPE '|' OR FULLPATH LIKE @nameObshVl ESCAPE '|') THEN 'BIM'
                        WHEN (FULLPATH LIKE @ovvk ESCAPE '|') THEN 'ОВиК, ВК'
                        WHEN (FULLPATH LIKE @eomss ESCAPE '|') THEN 'ЭОМ, СС'

                        WHEN (FULLPATH LIKE @ar1 ESCAPE '|' OR FULLPATH LIKE @ar2 ESCAPE '|') THEN 'АР'
                        WHEN (FULLPATH LIKE @kr1 ESCAPE '|' OR FULLPATH LIKE @kr2 ESCAPE '|') THEN 'КР'
                        WHEN (FULLPATH LIKE @ov1 ESCAPE '|' OR FULLPATH LIKE @ov2 ESCAPE '|' OR FULLPATH LIKE @ov3 ESCAPE '|') THEN 'ОВиК'
                        WHEN (FULLPATH LIKE @vk1 ESCAPE '|' OR FULLPATH LIKE @vk2 ESCAPE '|') THEN 'ВК'
                        WHEN (FULLPATH LIKE @eom1 ESCAPE '|' OR FULLPATH LIKE @eom2 ESCAPE '|') THEN 'ЭОМ'
                        WHEN (FULLPATH LIKE @ss1 ESCAPE '|' OR FULLPATH LIKE @ss2 ESCAPE '|') THEN 'СС'
                        ELSE DEPARTAMENT
                      END
                WHERE 
                  DEPARTAMENT IS NULL
                  AND STATUS IN ('NEW','UPDATE','OK')
                  AND (
                    -- Новые правила
                    FULLPATH LIKE @bimTemplates ESCAPE '|' OR FULLPATH LIKE @nameObshVl ESCAPE '|' OR
                    FULLPATH LIKE @ovvk ESCAPE '|' OR
                    FULLPATH LIKE @eomss ESCAPE '|' OR

                    -- Старые правила
                    FULLPATH LIKE @ar1 ESCAPE '|' OR FULLPATH LIKE @ar2 ESCAPE '|' OR
                    FULLPATH LIKE @kr1 ESCAPE '|' OR FULLPATH LIKE @kr2 ESCAPE '|' OR
                    FULLPATH LIKE @ov1 ESCAPE '|' OR FULLPATH LIKE @ov2 ESCAPE '|' OR FULLPATH LIKE @ov3 ESCAPE '|' OR
                    FULLPATH LIKE @vk1 ESCAPE '|' OR FULLPATH LIKE @vk2 ESCAPE '|' OR
                    FULLPATH LIKE @eom1 ESCAPE '|' OR FULLPATH LIKE @eom2 ESCAPE '|' OR
                    FULLPATH LIKE @ss1 ESCAPE '|' OR FULLPATH LIKE @ss2 ESCAPE '|'
                  );";


                    cmd.Parameters.AddWithValue("@bimTemplates", Like(@"X:\BIM\3_Семейства\7_Шаблоны семейств"));
                    cmd.Parameters.AddWithValue("@nameObshVl", ContainsLike("_ОбщВл_"));
                    cmd.Parameters.AddWithValue("@ovvk", Like(@"X:\BIM\3_Семейства\0_Общие семейства\2_ИОС\ОВВК"));
                    cmd.Parameters.AddWithValue("@eomss", Like(@"X:\BIM\3_Семейства\0_Общие семейства\2_ИОС\ЭОМСС"));
                    cmd.Parameters.AddWithValue("@ar1", Like(@"X:\BIM\3_Семейства\1_АР"));
                    cmd.Parameters.AddWithValue("@ar2", Like(@"X:\BIM\3_Семейства\8_Библиотека семейств Самолета\01_АР"));
                    cmd.Parameters.AddWithValue("@kr1", Like(@"X:\BIM\3_Семейства\2_КР"));
                    cmd.Parameters.AddWithValue("@kr2", Like(@"X:\BIM\3_Семейства\8_Библиотека семейств Самолета\02_КР"));
                    cmd.Parameters.AddWithValue("@ov1", Like(@"X:\BIM\3_Семейства\4_ОВиК"));
                    cmd.Parameters.AddWithValue("@ov2", Like(@"X:\BIM\3_Семейства\8_Библиотека семейств Самолета\08_ТМ"));
                    cmd.Parameters.AddWithValue("@ov3", Like(@"X:\BIM\3_Семейства\8_Библиотека семейств Самолета\03_ОВ"));
                    cmd.Parameters.AddWithValue("@vk1", Like(@"X:\BIM\3_Семейства\3_ВК"));
                    cmd.Parameters.AddWithValue("@vk2", Like(@"X:\BIM\3_Семейства\8_Библиотека семейств Самолета\04_ВК"));
                    cmd.Parameters.AddWithValue("@eom1", Like(@"X:\BIM\3_Семейства\6_ЭОМ"));
                    cmd.Parameters.AddWithValue("@eom2", Like(@"X:\BIM\3_Семейства\8_Библиотека семейств Самолета\05_ЭОМ"));
                    cmd.Parameters.AddWithValue("@ss1", Like(@"X:\BIM\3_Семейства\5_СС"));
                    cmd.Parameters.AddWithValue("@ss2", Like(@"X:\BIM\3_Семейства\8_Библиотека семейств Самолета\06_СС"));

                    int affected = cmd.ExecuteNonQuery();
                    tx.Commit();
                    return affected;
                }
            }
        }

        /// --- Обновлеие IMAGE в FamilyManager
        // Запись данных в БД (FamilyManager). STATUS, IMAGE
        private static int UpdateImagesByPath(string dbPath, Action<string> log = null, int maxSide = 200, int maxRecords = 4000)
        {
            if (string.IsNullOrWhiteSpace(dbPath) || !File.Exists(dbPath))
                throw new FileNotFoundException("Не найден файл БД", dbPath);

            if (maxSide < 1) maxSide = 1;
            if (maxRecords < 1) maxRecords = 1;

            int updated = 0;
            int processed = 0;

            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();

                using (var cmdSel = new SQLiteCommand(@"
                    SELECT ID, FULLPATH
                    FROM FamilyManager
                    WHERE (STATUS IN ('NEW','UPDATE','OK'))
                      AND (IMAGE IS NULL OR length(IMAGE)=0)
                    ORDER BY ID
                    LIMIT @lim;", conn))
                {
                    cmdSel.Parameters.AddWithValue("@lim", maxRecords);

                    using (var reader = cmdSel.ExecuteReader(System.Data.CommandBehavior.SequentialAccess))
                    {
                        while (reader.Read())
                        {
                            if (processed >= maxRecords)
                                break;

                            long id = SafeGetInt64(reader, 0);
                            string fullPath = SafeGetString(reader, 1);

                            processed++;

                            if (!IsValidRfaOrRvtPath(fullPath))
                                continue;

                            byte[] PicBytes = null;

                            try
                            {
                                using (Bitmap bmp = GetShellThumbnail(fullPath, 256))
                                {
                                    if (bmp != null)
                                    {
                                        using (Bitmap reduced = ResizeKeepingAspect(bmp, maxSide))
                                        {
                                            PicBytes = ToJpegBytes(reduced, 85L);
                                        }
                                    }
                                }
                            }
                            catch
                            {
                            }

                            if (PicBytes == null || PicBytes.Length == 0)
                                continue;

                            using (var cmdUpd = new SQLiteCommand(@"
                                UPDATE FamilyManager
                                SET IMAGE = @img
                                WHERE ID = @id
                                  AND (IMAGE IS NULL OR length(IMAGE)=0);", conn))
                            {
                                cmdUpd.Parameters.Add("@img", System.Data.DbType.Binary).Value = PicBytes;
                                cmdUpd.Parameters.Add("@id", System.Data.DbType.Int64).Value = id;

                                int rows = cmdUpd.ExecuteNonQuery();
                                if (rows > 0)
                                    updated += rows;
                            }
                        }
                    }
                }
            }

            return updated;
        }

        // Проверка путей
        private static bool IsValidRfaOrRvtPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                return false;

            string ext = Path.GetExtension(path);
            return ext != null &&
                   (ext.Equals(".rfa", StringComparison.OrdinalIgnoreCase));
        }

        // Чтение и получения превью
        private static long SafeGetInt64(System.Data.IDataRecord reader, int ordinal)
            => reader.IsDBNull(ordinal) ? 0L : Convert.ToInt64(reader.GetValue(ordinal));

        private static string SafeGetString(System.Data.IDataRecord reader, int ordinal)
            => reader.IsDBNull(ordinal) ? null : reader.GetString(ordinal);

        // Ресайз изображения
        private static Bitmap ResizeKeepingAspect(Bitmap src, int maxSide)
        {
            double scale = Math.Min((double)maxSide / src.Width, (double)maxSide / src.Height);
            int w = Math.Max(1, (int)Math.Round(src.Width * scale));
            int h = Math.Max(1, (int)Math.Round(src.Height * scale));

            var dest = new Bitmap(w, h, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            dest.SetResolution(src.HorizontalResolution, src.VerticalResolution);

            using (var g = Graphics.FromImage(dest))
            {
                g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                g.DrawImage(src, new System.Drawing.Rectangle(0, 0, w, h));
            }
            return dest;
        }

        // Конвертация в BLOP формат
        private static byte[] ToJpegBytes(Bitmap bmp, long quality = 85L)
        {
            using (var ms = new MemoryStream())
            {
                var codec = ImageCodecInfo.GetImageDecoders()
                    .FirstOrDefault(c => c.FormatID == ImageFormat.Jpeg.Guid);

                var encParams = new EncoderParameters(1);
                encParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, quality);

                bmp.Save(ms, codec, encParams);
                return ms.ToArray();
            }
        }

        // Windows Shell Thumbnail (COM)
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void SHCreateItemFromParsingName(
            [MarshalAs(UnmanagedType.LPWStr)] string pszPath,
            IntPtr pbc,
            [MarshalAs(UnmanagedType.LPStruct)] Guid riid,
            out IShellItemImageFactory ppv);

        // IShellItemImageFactory GUID
        private static readonly Guid IID_IShellItemImageFactory = new Guid("bcc18b79-ba16-442f-80c4-8a59c30c463b");

        [ComImport]
        [Guid("bcc18b79-ba16-442f-80c4-8a59c30c463b")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellItemImageFactory
        {
            void GetImage(SIZE size, SIIGBF flags, out IntPtr phbm);
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct SIZE
        {
            public int cx;
            public int cy;
        }

        [Flags]
        private enum SIIGBF
        {
            SIIGBF_RESIZETOFIT = 0x00,
            SIIGBF_BIGGERSIZEOK = 0x01,
            SIIGBF_MEMORYONLY = 0x02,
            SIIGBF_ICONONLY = 0x04,
            SIIGBF_THUMBNAILONLY = 0x08,
            SIIGBF_INCACHEONLY = 0x10,
        }

        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        private static Bitmap GetShellThumbnail(string path, int requested)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                return null;

            SHCreateItemFromParsingName(path, IntPtr.Zero, IID_IShellItemImageFactory, out var factory);

            IntPtr hBmp = IntPtr.Zero;
            try
            {
                var sz = new SIZE { cx = requested, cy = requested };
                factory.GetImage(sz, SIIGBF.SIIGBF_THUMBNAILONLY | SIIGBF.SIIGBF_BIGGERSIZEOK, out hBmp);
                if (hBmp == IntPtr.Zero)
                {
                    factory.GetImage(sz, SIIGBF.SIIGBF_ICONONLY | SIIGBF.SIIGBF_BIGGERSIZEOK, out hBmp);
                    if (hBmp == IntPtr.Zero)
                        return null;
                }

                var bmp = System.Drawing.Image.FromHbitmap(hBmp);
                return (Bitmap)bmp;
            }
            finally
            {
                if (hBmp != IntPtr.Zero)
                    DeleteObject(hBmp);
                if (factory != null)
                    Marshal.ReleaseComObject(factory);
            }
        }

        // Интерфейс универсального отдела
        private IEnumerable<FamilyManagerRecord> GetUniversalFiltered(string depUi)
        {
            var depDb = DepForDb(depUi);
            IEnumerable<FamilyManagerRecord> all = _records ?? Enumerable.Empty<FamilyManagerRecord>();

            var filtered = all.Where(r =>
                !string.IsNullOrWhiteSpace(r.Departament) &&
                DepContains(r.Departament, depDb) &&
                HasStatusNewOrOk(r.Status));

            string q = (_isWatermarkActive ? null : _tbSearch?.Text)?.Trim();
            if (!string.IsNullOrEmpty(q))
            {
                filtered = filtered.Where(r => MatchesSearchByIndex(r, q));
            }

            if (string.Equals(depUi?.Trim(), "АР", StringComparison.OrdinalIgnoreCase)
                && _cbStage != null
                && _cbStage.SelectedValue is int stageId && stageId > 0)
            {
                filtered = filtered.Where(r => r.Stage == stageId);
            }

            if (_cbProject != null && _cbProject.SelectedValue is int projectId && projectId > 0)
            {
                filtered = filtered.Where(r => r.Project == projectId);
            }

            return filtered.OrderBy(r => SafeFileName(r.FullPath), StringComparer.OrdinalIgnoreCase);
        }

        // Интерфейс универсального отдела. Фильтр по статусу
        private static bool HasStatusNewOrOk(string status)
        {
            if (string.IsNullOrWhiteSpace(status)) return false;
            var s = status.Trim();
            return string.Equals(s, "NEW", StringComparison.OrdinalIgnoreCase)
                || string.Equals(s, "OK", StringComparison.OrdinalIgnoreCase);
        }

        // Интерфейс универсального отдела. Фильтр по отделу
        private static bool DepContains(string depCell, string depFilter)
        {
            if (string.IsNullOrWhiteSpace(depCell) || string.IsNullOrWhiteSpace(depFilter))
                return false;
            return depCell.IndexOf(depFilter, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        // Интерфейс для общих отделов. Общее
        private void RebuildUniversalContent()
        {
            if (_universalRootPanel == null) return;

            CaptureUniversalScroll();
            _universalRootPanel.Children.Clear();

            var depUi = _universalDepartment ?? GetCurrentDepartment();
            var items = GetUniversalFiltered(depUi).ToList();

            bool isSearching = !_isWatermarkActive && !string.IsNullOrWhiteSpace(_tbSearch?.Text);

            if (items.Count == 0)
            {
                _universalRootPanel.Children.Add(new TextBlock
                {
                    Text = "Нет элементов для отображения.",
                    Opacity = 0.7,
                    Margin = new Thickness(0, 6, 0, 0)
                });
                UpdateSelectionButtonsState();
                RestoreUniversalScroll();
                return;
            }

            var favorites = items.Where(IsFavorite).ToList();
            if (favorites.Count > 0)
            {
                var favPanel = BuildRowsPanel(favorites);
                var asm = this.GetType().Assembly;
                var favRes = asm.GetName().Name + ".Imagens.FamilyManager.favorites.png";
                var favIcon = GetEmbeddedIconCached(favRes);

                var favExp = CreateUniversalExpander(
                    "Избранное", favorites.Count,     
                    "cat:favorites",
                    favPanel, favIcon,
                    forceExpanded: isSearching ? true : (bool?)null,
                    persistState: !isSearching,
                    iconSize: 32
                );
                _universalRootPanel.Children.Add(favExp);
            }

            var unassigned = new List<FamilyManagerRecord>();
            var byCategory = new Dictionary<int, List<FamilyManagerRecord>>();

            foreach (var r in items)
            {
                if (IsUnassigned(r))
                    unassigned.Add(r);
                else
                {
                    if (!byCategory.TryGetValue(r.Category, out var list))
                    {
                        list = new List<FamilyManagerRecord>();
                        byCategory[r.Category] = list;
                    }
                    list.Add(r);
                }
            }

            if (unassigned.Count > 0)
            {
                var unassignedPanel = BuildRowsPanel(unassigned);
                var asm = this.GetType().Assembly;
                var ngRes = asm.GetName().Name + ".Imagens.FamilyManager.nongroup.png";
                var unIcon = GetEmbeddedIconCached(ngRes);
                var unHeader = $"Не назначено ({unassigned.Count})";
                var exp = CreateUniversalExpander(
                    "Не назначено", unassigned.Count,
                    "cat:-1",
                    unassignedPanel, unIcon,
                    forceExpanded: isSearching ? true : (bool?)null,
                    persistState: !isSearching,
                    iconSize: 32
                );
                _universalRootPanel.Children.Add(exp);
            }

            foreach (var kv in byCategory)
            {
                int catId = kv.Key;
                var catRecs = kv.Value;

                TryGetCategoryInfo(catId, out string catName, out var subMap);

                var atRoot = new List<FamilyManagerRecord>();
                var bySub = new Dictionary<int, List<FamilyManagerRecord>>();

                foreach (var r in catRecs)
                {
                    if (r.SubCategory > 0 && subMap.TryGetValue(r.SubCategory, out var _))
                    {
                        if (!bySub.TryGetValue(r.SubCategory, out var lst))
                        {
                            lst = new List<FamilyManagerRecord>();
                            bySub[r.SubCategory] = lst;
                        }
                        lst.Add(r);
                    }
                    else
                    {
                        atRoot.Add(r);
                    }
                }

                var catStack = new StackPanel { Margin = new Thickness(0) };

                if (bySub.Count > 0)
                {
                    var orderedSubs = bySub
                        .Select(s =>
                        {
                            string subName = subMap.TryGetValue(s.Key, out var sn) && !string.IsNullOrWhiteSpace(sn)
                                ? sn
                                : $"Подкатегория {s.Key}";
                            int count = s.Value?.Count ?? 0;
                            return (subId: s.Key, subName, recs: s.Value, count);
                        })
                        .OrderBy(x => x.subName, StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    foreach (var (subId, subName, recs, count) in orderedSubs)
                    {
                        if (recs == null || recs.Count == 0) continue;

                        var subPanel = BuildRowsPanel(recs);
                        string subHeader = $"{subName} ({count})";
                        var subIcon = GetCategoryHeaderIconCached(catId, subId);
                        var subExp = CreateUniversalExpander(
                            subName, count,
                            $"cat:{catId}:sub:{subId}",
                            subPanel, subIcon,
                            forceExpanded: isSearching ? true : (bool?)null,
                            persistState: !isSearching,
                            iconSize: 32 
                        );
                        subExp.Margin = new Thickness(12, 4, 0, 0);
                        catStack.Children.Add(subExp);
                    }
                }


                if (atRoot.Count > 0)
                {
                    catStack.Children.Add(BuildRowsPanel(atRoot));
                }

                if (catStack.Children.Count > 0)
                {
                    var totalCount = catRecs.Count;
                    var catIcon = GetCategoryHeaderIconCached(catId, 0);
                    var catExp = CreateUniversalExpander(
                        catName, totalCount,
                        $"cat:{catId}",
                        catStack, catIcon,
                        forceExpanded: isSearching ? true : (bool?)null,
                        persistState: !isSearching,
                        iconSize: 32 
                    );
                    _universalRootPanel.Children.Add(catExp);
                }
            }

            var allExps = _universalRootPanel.Children.OfType<Expander>().ToList();

            var fav = allExps.Where(e => (e.Tag as string) == "cat:favorites").ToList();
            var un = allExps.Where(e => (e.Tag as string) == "cat:-1").ToList();
            var cats = allExps.Where(e =>
            {
                var tag = e.Tag as string;
                return tag != "cat:favorites" && tag != "cat:-1";
            })
            .OrderBy(e =>
            {
                var sp = e.Header as StackPanel;
                if (sp != null)
                {
                    var key = sp.Tag as string;
                    if (!string.IsNullOrEmpty(key)) return key;
                    var tb = sp.Children.OfType<TextBlock>().FirstOrDefault();
                    if (tb != null && tb.Tag is string tkey && !string.IsNullOrEmpty(tkey)) return tkey;
                }
                return ""; 
            }, StringComparer.OrdinalIgnoreCase)
            .ToList();

            _universalRootPanel.Children.Clear();
            foreach (var e in fav) _universalRootPanel.Children.Add(e);
            foreach (var e in un) _universalRootPanel.Children.Add(e);
            foreach (var e in cats) _universalRootPanel.Children.Add(e);

            UpdateSelectionButtonsState();
            RestoreUniversalScroll();
        }

        // Интерфейс универсального отдела. Биндер для ComboBox - Cтадия
        private void BindStageCombo_DefaultId1(ComboBox cb)
        {
            var items = new List<KeyValuePair<int, string>>
            {
                new KeyValuePair<int, string>(-1, "Для всех стадий (без фильтра)")
            };

            items.AddRange(
                _stagesById
                    .Where(kv =>
                        !string.Equals(kv.Value, "Для всех стадий", StringComparison.OrdinalIgnoreCase))
                    .OrderBy(kv => kv.Key)
                    .Select(kv => new KeyValuePair<int, string>(kv.Key, kv.Value))
            );

            cb.ItemsSource = items;
            cb.DisplayMemberPath = "Value";
            cb.SelectedValuePath = "Key";
            cb.SelectedValue = -1; 
        }

        // Интерфейс универсального отдела. Биндер для ComboBox - Проект
        private void BindProjectCombo_DefaultId1(ComboBox cb)
        {
            var items = new List<KeyValuePair<int, string>>
            {
                new KeyValuePair<int, string>(-1, "Для всех проектов (без фильтра)")
            };

            items.AddRange(
                _projectsById
                    .Where(kv => !string.Equals(kv.Value, "Для всех проектов", StringComparison.OrdinalIgnoreCase))
                    .OrderBy(kv => kv.Key)
                    .Select(kv => new KeyValuePair<int, string>(kv.Key, kv.Value))
            );

            cb.ItemsSource = items;
            cb.DisplayMemberPath = "Value";
            cb.SelectedValuePath = "Key";
            cb.SelectedValue = -1; 
        }

        // Интерфейс универсального отдела. Строковое представление
        private FrameworkElement CreateUniversalRow(FamilyManagerRecord rec)
        {
            var btn = new Button
            {
                Padding = new Thickness(0),
                Background = System.Windows.Media.Brushes.Transparent,
                BorderThickness = new Thickness(0),
                HorizontalContentAlignment = HorizontalAlignment.Stretch,
                VerticalContentAlignment = VerticalAlignment.Stretch,
                Tag = rec.ID
            };

            if (_currentSubDep == "BIM")
            {
                btn.Click += OnRowMainClick;
            }
            else
            {
                btn.Click += OnUniversalRowClick;
            }

            var border = new Border
            {
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(8),
                Margin = new Thickness(0, 4, 0, 2),
                BorderBrush = System.Windows.Media.Brushes.LightGray,
                Background = System.Windows.Media.Brushes.White
            };
            btn.Content = border;

            var grid = new Grid
            {
                Margin = new Thickness(0)
            };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto }); 
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            border.Child = grid;

            var cb = new CheckBox
            {
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 8, 0),
                Tag = rec
            };

            cb.Click += (s, e) =>
            {
                e.Handled = true; 
            };

            var deptKey = GetDeptKey();
            var set = GetOrCreateSelectionSet(deptKey);
            cb.IsChecked = set.Contains(rec.ID);

            cb.Checked += (s, e) =>
            {
                var r = (s as CheckBox)?.Tag as FamilyManagerRecord;
                if (r == null) return;
                GetOrCreateSelectionSet(GetDeptKey()).Add(r.ID);
                UpdateSelectionButtonsState();
                RefreshScenario(); 
                e.Handled = true;
            };
            cb.Unchecked += (s, e) =>
            {
                var r = (s as CheckBox)?.Tag as FamilyManagerRecord;
                if (r == null) return;
                GetOrCreateSelectionSet(GetDeptKey()).Remove(r.ID);
                UpdateSelectionButtonsState();
                RefreshScenario(); 
                e.Handled = true;
            };

            Grid.SetColumn(cb, 0);
            grid.Children.Add(cb);

            var img = new System.Windows.Controls.Image
            {
                Width = 48,
                Height = 48,
                Stretch = System.Windows.Media.Stretch.UniformToFill,
                Margin = new Thickness(0, 0, 8, 0),
                SnapsToDevicePixels = true
            };
            if (rec.ImageBytes != null && rec.ImageBytes.Length > 0)
            {
                try
                {
                    var bi = new BitmapImage();
                    using (var ms = new MemoryStream(rec.ImageBytes))
                    {
                        bi.BeginInit();
                        bi.CacheOption = BitmapCacheOption.OnLoad;
                        bi.StreamSource = ms;
                        bi.EndInit();
                        bi.Freeze();
                    }
                    img.Source = bi;
                }
                catch
                {
                }
            }
            Grid.SetColumn(img, 1);
            grid.Children.Add(img);

            var spText = new StackPanel
            {
                Orientation = Orientation.Vertical,
                VerticalAlignment = VerticalAlignment.Center
            };

            var nameRun = new TextBlock
            {
                Text = SafeFileName(rec.FullPath),
                FontWeight = FontWeights.SemiBold,
                TextTrimming = TextTrimming.CharacterEllipsis
            };

            var subRun = new TextBlock
            {
                Text = ResolveCategoryName(rec.Category, rec.SubCategory),
                Opacity = 0.7,
                FontStyle = FontStyles.Italic,
                Margin = new Thickness(0, 2, 0, 0),
                TextTrimming = TextTrimming.CharacterEllipsis
            };

            spText.Children.Add(nameRun);
            spText.Children.Add(subRun);

            Grid.SetColumn(spText, 2);
            grid.Children.Add(spText);

            ToolTipService.SetToolTip(btn, BuildFamilyTooltip(rec));

            return btn;
        }

        // Интерфейс универсального отдела. Преобразование категорий
        private string ResolveCategoryName(int categoryId, int subCategoryId)
        {
            if (categoryId == 1 && subCategoryId == 0)
                return "Категория не задана";

            var catPair = _categoriesById.Keys.FirstOrDefault(k => k.Id == categoryId);

            if (catPair != default)
            {
                var subMap = _categoriesById[catPair];

                if (subCategoryId > 0 && subMap.TryGetValue(subCategoryId, out var subName) && !string.IsNullOrWhiteSpace(subName))
                    return subName;

                if (subCategoryId == 0 && !string.IsNullOrWhiteSpace(catPair.Name))
                    return catPair.Name;
            }

            return subCategoryId == 0 ? "Корневая директория" : "Без категории";
        }

        // Интерфейс универсального отдела. Заглушка для кнопки действия
        private void OnUniversalRowClick(object sender, RoutedEventArgs e)
        {
            var idObj = (sender as System.Windows.Controls.Button)?.Tag;
            var idText = idObj?.ToString() ?? "null";
            OpenFamilyUserLoop(idText);
        }

        // Интерфейс универсального отдела. Выделение элементов
        private HashSet<int> GetOrCreateSelectionSet(string deptKey)
        {
            if (!_selectedByDept.TryGetValue(deptKey, out var set))
            {
                set = new HashSet<int>();
                _selectedByDept[deptKey] = set;
            }
            return set;
        }

        // Интерфейс универсального отдела. Сброс элементов
        private void ClearCurrentDeptSelection()
        {
            var key = GetDeptKey();
            if (string.IsNullOrEmpty(key)) return;
            if (_selectedByDept.ContainsKey(key))
                _selectedByDept[key].Clear();
        }

        // Интерфейс универсального отдела. Статус кнопок
        private void UpdateSelectionButtonsState()
        {
            if (_btnOpenInRevit == null || _btnLoadIntoProject == null) return;
            var hasAny = GetOrCreateSelectionSet(GetDeptKey()).Count > 0;
            _btnOpenInRevit.IsEnabled = hasAny;
            _btnLoadIntoProject.IsEnabled = hasAny;
        }

        // Интерфейс универсального отдела. Все семейства
        private List<FamilyManagerRecord> GetSelectedRecordsForCurrentDept()
        {
            var key = GetDeptKey();
            var set = GetOrCreateSelectionSet(key);
            if (set.Count == 0) return new List<FamilyManagerRecord>();

            IEnumerable<FamilyManagerRecord> pool = _records ?? Enumerable.Empty<FamilyManagerRecord>();
            var depUi = _universalDepartment ?? GetCurrentDepartment();
            var depDb = DepForDb(depUi);
            pool = pool.Where(r =>
                !string.IsNullOrWhiteSpace(r.Departament) &&
                DepContains(r.Departament, depDb) &&
                HasStatusNewOrOk(r.Status));

            return pool.Where(r => set.Contains(r.ID)).ToList();
        }

        // Определяем, что запись "не назначена" по категории
        private static bool IsUnassigned(FamilyManagerRecord r)
        {
            if (r == null) return true;
            if (r.Category == 0) return true;
            if (r.Category == 1 && r.SubCategory == 0) return true;
            return false;
        }

        // Получаем имя категории и карту подкатегорий по categoryId
        private bool TryGetCategoryInfo(int categoryId, out string categoryName, out Dictionary<int, string> subMap)
        {
            var pair = _categoriesById.Keys.FirstOrDefault(k => k.Id == categoryId);
            if (pair != default)
            {
                categoryName = string.IsNullOrWhiteSpace(pair.Name) ? $"Категория {categoryId}" : pair.Name;
                subMap = _categoriesById[pair] ?? new Dictionary<int, string>();
                return true;
            }
            categoryName = $"Категория {categoryId}";
            subMap = new Dictionary<int, string>();
            return false;
        }

        // Создаёт контейнер со строками элементов, отсортированными: выбранные — наверх, потом по имени
        private System.Windows.Controls.Panel BuildRowsPanel(IEnumerable<FamilyManagerRecord> recs)
        {
            var sp = new StackPanel { Margin = new Thickness(8, 4, 0, 6) };
            var set = GetOrCreateSelectionSet(GetDeptKey());

            foreach (var r in recs
                     .OrderByDescending(x => set.Contains(x.ID))
                     .ThenBy(x => SafeFileName(x.FullPath), StringComparer.OrdinalIgnoreCase))
            {
                sp.Children.Add(CreateUniversalRow(r));
            }
            return sp;
        }

        // Перед очисткой — запоминаем позицию скролла
        private void CaptureUniversalScroll()
        {
            if (_univScrollViewer != null)
                _univScrollOffset = _univScrollViewer.VerticalOffset;
        }
        // После перестройки — восстанавливаем скролл
        private void RestoreUniversalScroll()
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (_univScrollViewer != null)
                    _univScrollViewer.ScrollToVerticalOffset(_univScrollOffset);
            }), System.Windows.Threading.DispatcherPriority.Loaded);
        }

        // Шапка экспандера: [стрелка] [картинка] [текст]
        private object BuildExpanderHeader(string mainText, int? count, BitmapImage icon, double iconSize = 32)
        {
            var sp = new StackPanel
            {
                Tag = mainText,
                Orientation = Orientation.Horizontal,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 2, 0, 2)
            };

            if (icon != null)
            {
                sp.Children.Add(new System.Windows.Controls.Image
                {
                    Source = icon,
                    Width = iconSize,
                    Height = iconSize,
                    Margin = new Thickness(6, 0, 8, 0),
                    SnapsToDevicePixels = true
                });
            }

            var tb = new TextBlock
            {
                Tag = mainText,
                VerticalAlignment = VerticalAlignment.Center,
                FontSize = 13,
                TextTrimming = TextTrimming.CharacterEllipsis
            };
            tb.Inlines.Add(new Run(mainText) { FontWeight = FontWeights.SemiBold });
            if (count.HasValue)
            {
                tb.Inlines.Add(new Run(" (" + count.Value.ToString() + ")")
                {
                    FontWeight = FontWeights.Normal,
                });
            }

            sp.Children.Add(tb);
            return sp;
        }

        // Экспандер с запоминанием состояния
        private Expander CreateUniversalExpander(string mainText, int? count, string stateKey, UIElement content, BitmapImage headerIcon = null, bool? forceExpanded = null, bool persistState = true, double iconSize = 32)
        {
            bool isExpanded = _univExpandState.TryGetValue(stateKey, out var saved) ? saved : false;
            if (forceExpanded.HasValue) isExpanded = forceExpanded.Value;

            var exp = new Expander
            {
                Header = BuildExpanderHeader(mainText, count, headerIcon, iconSize),
                IsExpanded = isExpanded,
                Margin = new Thickness(0, 6, 0, 0),
                Content = content,
                Tag = stateKey
            };

            if (persistState)
            {
                exp.Expanded += (s, e) => _univExpandState[stateKey] = true;
                exp.Collapsed += (s, e) => _univExpandState[stateKey] = false;
            }
            return exp;
        }

        // Интерфейс универсального отдела. Открыть в Revit
        private void OnOpenInRevitClick(object sender, RoutedEventArgs e)
        {
            var list = GetSelectedRecordsForCurrentDept();
            if (list.Count == 0) return;

            var names = list.Select(r => SafeFileName(r.FullPath)).ToList();

            var msg = "Сейчас будут открыты в Revit следующие семейства:\n" +
                      string.Join("\n", names);
            var result = System.Windows.MessageBox.Show(
                msg, "Подтверждение", MessageBoxButton.OKCancel, MessageBoxImage.Question);

            if (result != MessageBoxResult.OK) return;

            if (_uiapp == null)
            {
                TaskDialog.Show("Ошибка", "Не удалось установить UIApplication.");
                return;
            }

            int ok = 0, err = 0;
            foreach (var r in list)
            {
                try
                {
                    if (!System.IO.File.Exists(r.FullPath)) { err++; continue; }
                    _uiapp.OpenAndActivateDocument(r.FullPath);
                    ok++;
                }
                catch { err++; }
            }
        }

        // Интерфейс универсального отдела. Открыть в Revit (загрузить в проект)
        private void OnLoadIntoProjectClick(object sender, RoutedEventArgs e)
        {
            var list = GetSelectedRecordsForCurrentDept();
            if (list.Count == 0) return;

            var names = list.Select(r => SafeFileName(r.FullPath)).ToList();

            var msg = "Сейчас будут загружены в активный проект следующие семейства:\n" +
                      string.Join("\n", names);
            var result = System.Windows.MessageBox.Show(
                msg, "Подтверждение", MessageBoxButton.OKCancel, MessageBoxImage.Question);

            if (result != MessageBoxResult.OK) return;

            if (_uiapp == null)
            {
                TaskDialog.Show("Ошибка", "Не удалось установить UIApplication.");
                return;
            }

            ExternalEventsHost.EnsureCreated();
            var handler = ExternalEventsHost.LoadFamilyHandler;
            var evnt = ExternalEventsHost.LoadFamilyEvent;

            handler.BatchPaths = list
                .Select(r => r.FullPath)
                .Where(p => !string.IsNullOrWhiteSpace(p) && File.Exists(p))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (handler.BatchPaths.Count == 0)
            {
                TaskDialog.Show("Загрузка семейств", "Нет валидных путей для загрузки.");
                return;
            }

            evnt.Raise();
        }

        // Тултип для карточки семейства
        private ToolTip BuildFamilyTooltip(FamilyManagerRecord rec)
        {
            string name = SafeFileName(rec.FullPath);
            string dept = rec.Departament ?? "—";
            string stage = _stagesById.TryGetValue(rec.Stage, out var stageName) ? stageName : "—";
            string project = _projectsById.TryGetValue(rec.Project, out var projectName) ? projectName : "—";

            string category = "—";
            string subCategory = "—";

            var catPair = _categoriesById.Keys.FirstOrDefault(k => k.Id == rec.Category);
            if (catPair != default)
            {
                category = !string.IsNullOrWhiteSpace(catPair.Name) ? catPair.Name : "—";

                var subMap = _categoriesById[catPair];
                if (rec.SubCategory > 0 && subMap.TryGetValue(rec.SubCategory, out var subName) && !string.IsNullOrWhiteSpace(subName))
                    subCategory = subName;
            }

            var imgCtrl = new System.Windows.Controls.Image
            {
                Width = 170,
                Height = 170,
                Stretch = System.Windows.Media.Stretch.UniformToFill,
                Margin = new Thickness(0, 0, 12, 0),
                SnapsToDevicePixels = true
            };
            var bi = CreateBitmapImageFromBytes(rec.ImageBytes);
            if (bi != null)
                imgCtrl.Source = bi;

            var textStack = new StackPanel { Orientation = Orientation.Vertical, VerticalAlignment = VerticalAlignment.Top };

            var title = new TextBlock
            {
                Text = name,
                FontWeight = FontWeights.Bold,
                FontSize = 14,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            };
            textStack.Children.Add(title);

            void AddRow(string label, string value)
            {
                var sp = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 3) };
                sp.Children.Add(new TextBlock
                {
                    Text = $"{label}: ",
                    FontWeight = FontWeights.Bold,
                    Width = 100
                });
                sp.Children.Add(new TextBlock
                {
                    Text = value,
                    TextWrapping = TextWrapping.Wrap
                });
                textStack.Children.Add(sp);
            }

            AddRow("Отдел", dept);
            AddRow("Стадия", stage);
            AddRow("Проект", project);
            AddRow("Категория", category);
            AddRow("Подкатегория", subCategory);

            var grid = new Grid { Margin = new Thickness(0) };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            Grid.SetColumn(imgCtrl, 0);
            Grid.SetColumn(textStack, 1);
            grid.Children.Add(imgCtrl);
            grid.Children.Add(textStack);

            var border = new Border
            {
                Padding = new Thickness(10),
                Background = System.Windows.SystemColors.InfoBrush,
                BorderBrush = System.Windows.SystemColors.ActiveBorderBrush,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                Child = grid,
                MaxWidth = 580
            };

            var tip = new ToolTip { Content = border };
            ToolTipService.SetShowDuration(border, 30000);
            ToolTipService.SetInitialShowDelay(border, 200);
            ToolTipService.SetBetweenShowDelay(border, 100);
            return tip;
        }

        // Тултип. Обработка изображения
        private static BitmapImage CreateBitmapImageFromBytes(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0) return null;
            try
            {
                using (var ms = new MemoryStream(bytes))
                {
                    var bi = new BitmapImage();
                    bi.BeginInit();
                    bi.CacheOption = BitmapCacheOption.OnLoad;
                    bi.StreamSource = ms;
                    bi.EndInit();
                    bi.Freeze();
                    return bi;
                }
            }
            catch { return null; }
        }

        // Избраное. Ссылка на документ
        private static string GetFavoritesPath()
        {
            try
            {
                var localDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                if (string.IsNullOrWhiteSpace(localDir))
                    return null;
                return Path.Combine(localDir, "RevitFamilyManagerFavorites.txt");
            }
            catch
            {
                return null;
            }
        }

        // Избранное. Загрузка первичная
        private bool LoadFavoritesFromFile()
        {
            try
            {
                if (_favoritesPath == null)
                    _favoritesPath = GetFavoritesPath();

                if (string.IsNullOrWhiteSpace(_favoritesPath))
                    return false;

                if (!File.Exists(_favoritesPath))
                {
                    if (_favoriteIds.Count > 0)
                    {
                        _favoriteIds.Clear();
                        return true;
                    }
                    return false;
                }

                var fi = new FileInfo(_favoritesPath);
                if (fi.LastWriteTimeUtc <= _favoritesLastWrite)
                    return false;

                var text = File.ReadAllText(_favoritesPath, Encoding.UTF8);

                var newSet = new HashSet<int>();
                foreach (var token in text.Split(new[] { ',', ';', '\n', '\r', '\t', ' ' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    int id;
                    if (int.TryParse(token.Trim(), out id) && id > 0)
                        newSet.Add(id);
                }

                _favoriteIds.Clear();
                foreach (var id in newSet) _favoriteIds.Add(id);
                _favoritesLastWrite = fi.LastWriteTimeUtc;
                return true;
            }
            catch
            {
                return false;
            }
        }

        // Избранное. Загрузка повторная
        private void StartFavoritesWatcher()
        {
            try
            {
                if (_favoritesPath == null)
                    _favoritesPath = GetFavoritesPath();

                if (string.IsNullOrWhiteSpace(_favoritesPath)) return;

                if (_favoritesWatcher != null)
                {
                    _favoritesWatcher.EnableRaisingEvents = false;
                    _favoritesWatcher.Dispose();
                    _favoritesWatcher = null;
                }

                var dir = Path.GetDirectoryName(_favoritesPath);
                var file = Path.GetFileName(_favoritesPath);
                if (string.IsNullOrWhiteSpace(dir) || string.IsNullOrWhiteSpace(file)) return;

                _favoritesWatcher = new FileSystemWatcher(dir, file)
                {
                    NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.CreationTime | NotifyFilters.Size | NotifyFilters.FileName,
                    IncludeSubdirectories = false,
                    EnableRaisingEvents = true
                };

                FileSystemEventHandler onChange = (s, e) =>
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        if (LoadFavoritesFromFile())
                            RefreshScenario();
                    }));
                };

                _favoritesWatcher.Changed += onChange;
                _favoritesWatcher.Created += onChange;
                _favoritesWatcher.Renamed += (s, e) => onChange(s, e);
                _favoritesWatcher.Deleted += onChange;
            }
            catch { }
        }

        // Данные БД. Картинки категорий/подкатегорий (таблица Category_PIC)
        private void LoadCategoryPics(string dbPath)
        {
            _categoryPics.Clear();
            if (string.IsNullOrWhiteSpace(dbPath) || !File.Exists(dbPath))
                return;

            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand(@"
            SELECT CAT_ID, SUBCAT_ID, PIC
            FROM Category_PIC
            WHERE CAT_ID IS NOT NULL;", conn))
                using (var rd = cmd.ExecuteReader(System.Data.CommandBehavior.SequentialAccess))
                {
                    while (rd.Read())
                    {
                        int catId = rd.IsDBNull(0) ? 0 : Convert.ToInt32(rd.GetValue(0));
                        int subId = rd.IsDBNull(1) ? 0 : Convert.ToInt32(rd.GetValue(1)); 
                        byte[] bytes = rd.IsDBNull(2) ? null : (byte[])rd[2];

                        if (catId > 0 && bytes != null && bytes.Length > 0)
                            _categoryPics[(catId, subId)] = bytes;
                    }
                }
            }
        }

        // Загрузка иконок. Embedded PNG → BitmapImage (с кешем)
        private BitmapImage GetEmbeddedIconCached(string resourceName)
        {
            BitmapImage cached;
            if (_embeddedIconCache.TryGetValue(resourceName, out cached))
                return cached;

            try
            {
                var asm = this.GetType().Assembly;
                using (var s = asm.GetManifestResourceStream(resourceName))
                {
                    if (s == null) return null;
                    var bi = new BitmapImage();
                    bi.BeginInit();
                    bi.CacheOption = BitmapCacheOption.OnLoad;
                    bi.StreamSource = s;
                    bi.EndInit();
                    bi.Freeze();
                    _embeddedIconCache[resourceName] = bi;
                    return bi;
                }
            }
            catch { return null; }
        }

        // Загрузка иконок. Категории/подкатегории из БД
        private BitmapImage GetCategoryHeaderIconCached(int catId, int subId)
        {
            BitmapImage cached;
            if (_categoryIconCache.TryGetValue((catId, subId), out cached))
                return cached;

            byte[] bytes;
            if (_categoryPics.TryGetValue((catId, subId), out bytes) && bytes != null && bytes.Length > 0)
            {
                try
                {
                    var bi = CreateBitmapImageFromBytes(bytes);
                    if (bi != null)
                    {
                        _categoryIconCache[(catId, subId)] = bi;
                        return bi;
                    }
                }
                catch { }
            }
            _categoryIconCache[(catId, subId)] = null; 
            return null;
        }
    }
}