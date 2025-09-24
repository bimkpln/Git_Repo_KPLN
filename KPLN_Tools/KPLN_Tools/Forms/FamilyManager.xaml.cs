﻿using Autodesk.Revit.DB;
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
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Action = System.Action;
using Application = System.Windows.Application;
using Button = System.Windows.Controls.Button;
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

    // IExternalEventHandler. Загрузка семейства в проект
    internal class LoadFamilyHandler : IExternalEventHandler
    {
        public string FilePath;

        public void Execute(UIApplication app)
        {
            if (string.IsNullOrEmpty(FilePath) || !File.Exists(FilePath))
            {
                TaskDialog.Show("Ошибка", $"Файл не найден: {FilePath}");
                return;
            }

            UIDocument uidoc = app.ActiveUIDocument;
            if (uidoc == null)
            {
                TaskDialog.Show("Ошибка", "Нет активного проекта для загрузки семейства.");
                return;
            }

            Document doc = uidoc.Document;

            using (Transaction t = new Transaction(doc, "KPLN. Загрузить семейство"))
            {
                t.Start();
                if (doc.LoadFamily(FilePath, out Family fam))
                {
                    TaskDialog.Show("Загрузка семейства", $"Семейство «{fam?.Name}» загружено.");
                }
                else
                {
                    TaskDialog.Show("Загрузка семейства", "Не удалось загрузить семейство.");
                }
                t.Commit();
            }
        }

        public string GetName() => "KPLN.LoadFamilyHandler";
    }

    // IExternalEventHandler. Загрузка данных из семейств
    internal class BulkPagedUpdateHandler : IExternalEventHandler
    {
        public string DbPath;
        public int PageSize = 200;
        public volatile bool IsRunning;
        public int Selected, Updated, Skipped, Errors;
        public Action<BulkPagedUpdateHandler> Completed;

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

                                try
                                {
                                    pIdRead.Value = id;

                                    string status = null, full = null, dep = null;
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
                                    { Skipped++; continue; }

                                    if (string.IsNullOrWhiteSpace(dep)) { Skipped++; continue; }
                                    if (string.IsNullOrWhiteSpace(full) || !File.Exists(full))
                                    { Errors++; continue; }

                                    string json = FamilyManager.ReadImportInfoFromFamily(app, full, dep);
                                    if (string.IsNullOrWhiteSpace(json)) { Skipped++; continue; }

                                    pIdUpd.Value = id;
                                    pJsonUpd.Value = json;
                                    int rows = cmdUpd.ExecuteNonQuery();

                                    if (rows > 0) Updated++;
                                    else Errors++;
                                }
                                catch
                                {
                                    Errors++;
                                }
                            }

                            tx.Commit();
                        }
                    }
                }
            }
            catch
            {
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
        public string LastModifiedDate { get; set; }
        public int Category { get; set; }
        public int Project { get; set; }
        public int Stage { get; set; }
        public string Departament { get; set; }
        public string ImportInfo { get; set; }
    }

    // Док-панель
    public partial class FamilyManager : UserControl
    {
        private UIApplication _uiapp;

        private const string ITEM_ERROR = "ОШИБКА";
        private const string DB_PATH = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_FamilyManager.db";
        private const string RFA_ROOT = @"X:\BIM\3_Семейства";

        private TextBox _tbSearch;
        private FrameworkElement _scenarioContent;
        private StackPanel _bimRootPanel;

        string _currentSubDep;
        private List<FamilyManagerRecord> _records;

        private bool _depsTried = false;  
        private bool _depsLoaded = false;
        private Dictionary<int, string> _searchIndex = new Dictionary<int, string>();
        private DispatcherTimer _searchDebounceTimer;

        private Dictionary<string, bool> _bimExpandState = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        private double _bimScrollOffset = 0;
        private ScrollViewer _bimScrollViewer;

        public void SetUIApplication(UIApplication uiapp)
        {
            _uiapp = uiapp;
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
            string dep = GetCurrentDepartment();
            bool isError = dep == ITEM_ERROR;
            bool isBim = string.Equals(dep, "BIM", StringComparison.OrdinalIgnoreCase);

            CmbDepartment.IsEnabled = _currentSubDep == "BIM";
            BtnReload.IsEnabled = !(isError);
            BtnSettings.IsEnabled = isBim;
        }

        // Данные БД. Семейства
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
                            LastModifiedDate = reader.GetString(3),
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

        // XAML. Загрузка формы
        // Стартовый список отделов из конструктора (без БД)
        private void UserControl_Loaded(object sender, System.Windows.RoutedEventArgs e)
        {
            var items = new List<string>();
            if (!string.IsNullOrWhiteSpace(_currentSubDep))
            {
                items.Add(_currentSubDep);
            }
            else
            {
                items.Add(ITEM_ERROR);
            }

            CmbDepartment.ItemsSource = items;
            CmbDepartment.SelectedItem = items.First();

            UpdateUiState();
        }

        // XAML. Обновление статуса CmbDepartment
        private void CmbDepartment_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            MainArea.Child = null;

            var dep = GetCurrentDepartment();
            if (string.Equals(dep, "BIM", StringComparison.OrdinalIgnoreCase) && !_depsTried)
            {
                if (_depsTried) return;
                _depsTried = true;

                try
                {
                    var deps = LoadDepartments(DB_PATH);
                    if (deps.Count == 0)
                    {
                        CmbDepartment.ItemsSource = new List<string> { ITEM_ERROR };
                        CmbDepartment.SelectedItem = ITEM_ERROR;
                        _depsLoaded = false;
                        UpdateUiState();

                        TaskDialog.Show("Ошибка", "Ошибка чтения из БД");
                        return;
                    }

                    _depsLoaded = true;
                    var prev = CmbDepartment.SelectedItem?.ToString();
                    CmbDepartment.ItemsSource = deps;

                    if (!string.IsNullOrWhiteSpace(prev))
                    {
                        var match = deps.FirstOrDefault(d => string.Equals(d, prev, StringComparison.OrdinalIgnoreCase));
                        if (match != null)
                            CmbDepartment.SelectedItem = match;
                    }

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

            UpdateUiState();
        }

        // XAML. Загрузка данных по кнопке
        private void BtnReload_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (DepartmentExistsInDb(DB_PATH, GetCurrentDepartment()))
            {
                try
                {
                    _records = LoadFamilyManagerRecords(DB_PATH);
                    RebuildSearchIndex();
                    BuildMainArea(GetCurrentDepartment());
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

        // Индексация IMPORT_INFO. Собираем строковые/числовые значения из JSON
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

        // Индексация IMPORT_INFO. Строим текст для поиска: имя файла + все значения из IMPORT_INFO
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

        // Индексация IMPORT_INFO. Перестроение индекса для всех записей
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
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); 

            _tbSearch = new TextBox
            {
                MinWidth = 200,
                Height = 26,
                Margin = new Thickness(0, 0, 0, 8),
                Padding = new Thickness(6, 2, 6, 2),
                VerticalContentAlignment = VerticalAlignment.Center,
                ToolTip = "Поиск семейства по названию"
            };
            _tbSearch.TextChanged += OnSearchTextChanged;

            _searchDebounceTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(200) };
            _searchDebounceTimer.Tick += (s, e) =>
            {
                _searchDebounceTimer.Stop();
                RefreshScenario(); 
            };

            Grid.SetRow(_tbSearch, 0);
            root.Children.Add(_tbSearch);

            FrameworkElement scenarioUI = BuildScenarioUI(subDepartamentName, _records);

            Grid.SetRow(scenarioUI, 1);
            root.Children.Add(scenarioUI);

            _scenarioContent = scenarioUI;
            MainArea.Child = root;
        }

        // Обработчик собыйтий. Фильтр семейств по названию
        private void OnSearchTextChanged(object sender, TextChangedEventArgs e)
        {
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
        }

        // UI в зависимости от подразделения
        private FrameworkElement BuildScenarioUI(string dep, List<FamilyManagerRecord> records)
        {
            dep = dep?.Trim();

            if (string.Equals(dep, "BIM", StringComparison.OrdinalIgnoreCase))
            {
                var scroll = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
                _bimScrollViewer = scroll;
                _bimRootPanel = new StackPanel { Margin = new Thickness(0) };
                scroll.Content = _bimRootPanel;

                RebuildBimContent(); 
                return scroll;
            }
            else
            {
                return new StackPanel
                {
                    Children =
            {
                new TextBlock { Text = $"Подразделение {dep ?? "—"} — сюда придёт UI", Opacity = 0.7 }
            }
                };
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

            string q = _tbSearch?.Text?.Trim() ?? "";
            if (!string.IsNullOrEmpty(q))
            {
                string qNorm = q.ToUpperInvariant();

                all = all.Where(r =>
                {
                    string hay;
                    if (_searchIndex != null && _searchIndex.TryGetValue(r.ID, out hay) && hay != null)
                        return hay.Contains(qNorm);
                    return SafeFileName(r.FullPath)
                           .IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0;
                }).ToList();
            }

            Func<string, string> norm = s => (s ?? "").Trim().ToUpperInvariant();
            var catNew = all.Where(r => norm(r.Status) == "NEW").ToList();
            var catAbsentError = all.Where(r =>
            {
                var st = norm(r.Status);
                return st == "ABSENT" || st == "ERROR";
            }).ToList();

            var catNotDepartament = all.Where(r => r.Departament == null &&norm(r.Status) != "ABSENT" && norm(r.Status) != "ERROR" && norm(r.Status) != "IGNORE").ToList();
            var catOkWithBadMeta = all.Where(r =>(norm(r.Status) != "ABSENT" && norm(r.Status) != "ERROR" && norm(r.Status) != "IGNORE") && r.Category == 1).ToList();
            var catOkMissingImportOrImage = all.Where(r =>(norm(r.Status) != "ABSENT" && norm(r.Status) != "ERROR" && norm(r.Status) != "IGNORE") && (r.ImportInfo == null)).ToList();
            var catIgnored = all.Where(r => norm(r.Status) == "IGNORE").ToList();
            var catOkProcessed = all.Where(r => (norm(r.Status) == "OK") && r.Departament != null && r.Category != 1).ToList();

            _bimRootPanel.Children.Add(CreateCategoryExpander("НОВЫЕ СЕМЕЙСТВА", "bim_new.png", catNew, "Семейства, которые ранее были добавлены на диск и не были обработаны.\nДанные семейства не отображаются в списке у пользователей до изменения статуса на «OK» у семейства BIM-координатором.", key: "new"));
            _bimRootPanel.Children.Add(CreateCategoryExpander("ОШИБКИ / НЕ НАЙДЕН", "bim_error.png", catAbsentError, "Семейства, которые были удалены, или семейства, в которые параметры были экспортированы с ошибками.\nДанные семейства не отображаются в списке у пользователей до исправления ошибок BIM-координатором.", key: "errorabsent"));
            _bimRootPanel.Children.Add(CreateCategoryExpander("НЕТ ОТДЕЛА", "bim_error.png", catNotDepartament, "Семейства, которые не содержут информацию об отделе.\nДанные семейства не отображаются в списке у пользователей до указания отдела BIM-координатором.", key: "nodept"));
            _bimRootPanel.Children.Add(CreateCategoryExpander("НЕ УКАЗАНА КАТЕГОРИЯ", "bim_caution.png", catOkWithBadMeta, "Семейства, в которых не указан параметр КАТЕГОРИЯ.\nБез указания данного параметра семейство группируется в директории по-умолчанию.", key: "badmeta"));
            _bimRootPanel.Children.Add(CreateCategoryExpander("НЕТ СВОЙСТВ СЕМЕЙСТВА", "bim_caution.png", catOkMissingImportOrImage, "Семейства, в которых не указаны экспортируемые свойства семейства.\nБез указания данных параметров семейство не содержит описания о себе (из файла).", key: "missingimport"));           
            _bimRootPanel.Children.Add(CreateCategoryExpander("ИГНОРИРУЮТСЯ", "bim_ignore.png", catIgnored, "Семейства, помеченные к игнорированию. Не отображаются у пользователей.", key: "ignored"));
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

        // Логика открытия окна
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
                MessageBox.Show("Не удалось сформировать JSON: ошибка обработки параметров.", "Family Manager", MessageBoxButton.OK, MessageBoxImage.Information);
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
                    MessageBox.Show("Запись в БД не изменёна (возможно, запись не найдена).", "Family Manager",MessageBoxButton.OK, MessageBoxImage.Warning);
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
                    return "{\"NULL\":\"\"}"; 

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
        // IMPORT_INFO. --- АР ---
        private static string ExtractJson_AR(Document doc, Autodesk.Revit.DB.FamilyManager fm)
        {
            FamilyParameter pNote = fm.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            string note = FirstNonEmptyByTypes(doc, fm, pNote) ?? "";
            return "{"
                 + "\"Примечание\":\"" + JsonEscape(note) + "\""
                 + "}";
        }

        // IMPORT_INFO. --- КР ---
        private static string ExtractJson_KR(Document doc, Autodesk.Revit.DB.FamilyManager fm)
        {
            FamilyParameter pNote = fm.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            string note = FirstNonEmptyByTypes(doc, fm, pNote) ?? "";
            return "{"
                 + "\"Примечание\":\"" + JsonEscape(note) + "\""
                 + "}";
        }

        // IMPORT_INFO. --- ОВиК ---
        private static string ExtractJson_OViK(Document doc, Autodesk.Revit.DB.FamilyManager fm)
        {
            var pNote = fm.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            var pDescr = fm.get_Parameter(new Guid("f194bf60-b880-4217-b793-1e0c30dda5e9")); // КП_О_Наименование
            var pMark = fm.get_Parameter(new Guid("2204049c-d557-4dfc-8d70-13f19715e46d")); // КП_О_Марка
            var pManuf = fm.get_Parameter(new Guid("a8cdbf7b-d60a-485e-a520-447d2055f351")); // КП_О_Завод-изготовитель

            string note = FirstNonEmptyByTypes(doc, fm, pNote) ?? "";
            string descr = FirstNonEmptyByTypes(doc, fm, pDescr) ?? "";
            string mark = FirstNonEmptyByTypes(doc, fm, pMark) ?? "";
            string manuf = FirstNonEmptyByTypes(doc, fm, pManuf) ?? "";
           
            return "{"
                 + "\"Описание\":\"" + JsonEscape(descr) + "\","
                 + "\"Марка\":\"" + JsonEscape(mark) + "\","
                 + "\"Производитель\":\"" + JsonEscape(manuf) + "\","
                 + "\"Примечание\":\"" + JsonEscape(note) + "\""
                 + "}";
        }

        // IMPORT_INFO. --- ВК ---
        private static string ExtractJson_VK(Document doc, Autodesk.Revit.DB.FamilyManager fm)
        {
            var pNote = fm.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            var pDescr = fm.get_Parameter(new Guid("f194bf60-b880-4217-b793-1e0c30dda5e9")); // КП_О_Наименование
            var pMark = fm.get_Parameter(new Guid("2204049c-d557-4dfc-8d70-13f19715e46d")); // КП_О_Марка
            var pManuf = fm.get_Parameter(new Guid("a8cdbf7b-d60a-485e-a520-447d2055f351")); // КП_О_Завод-изготовитель

            string note = FirstNonEmptyByTypes(doc, fm, pNote) ?? "";
            string descr = FirstNonEmptyByTypes(doc, fm, pDescr) ?? "";
            string mark = FirstNonEmptyByTypes(doc, fm, pMark) ?? "";
            string manuf = FirstNonEmptyByTypes(doc, fm, pManuf) ?? "";

            return "{"
                 + "\"Описание\":\"" + JsonEscape(descr) + "\","
                 + "\"Марка\":\"" + JsonEscape(mark) + "\","
                 + "\"Производитель\":\"" + JsonEscape(manuf) + "\","
                 + "\"Примечание\":\"" + JsonEscape(note) + "\""
                 + "}";
        }

        // IMPORT_INFO. --- ЭОМ ---
        private static string ExtractJson_EOM(Document doc, Autodesk.Revit.DB.FamilyManager fm)
        {
            var pNote = fm.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            var pDescr = fm.get_Parameter(new Guid("f194bf60-b880-4217-b793-1e0c30dda5e9")); // КП_О_Наименование
            var pMark = fm.get_Parameter(new Guid("2204049c-d557-4dfc-8d70-13f19715e46d")); // КП_О_Марка
            var pManuf = fm.get_Parameter(new Guid("a8cdbf7b-d60a-485e-a520-447d2055f351")); // КП_О_Завод-изготовитель

            string note = FirstNonEmptyByTypes(doc, fm, pNote) ?? "";
            string descr = FirstNonEmptyByTypes(doc, fm, pDescr) ?? "";
            string mark = FirstNonEmptyByTypes(doc, fm, pMark) ?? "";
            string manuf = FirstNonEmptyByTypes(doc, fm, pManuf) ?? "";

            return "{"
                 + "\"Описание\":\"" + JsonEscape(descr) + "\","
                 + "\"Марка\":\"" + JsonEscape(mark) + "\","
                 + "\"Производитель\":\"" + JsonEscape(manuf) + "\","
                 + "\"Примечание\":\"" + JsonEscape(note) + "\""
                 + "}";
        }

        // IMPORT_INFO. --- СС ---
        private static string ExtractJson_SS(Document doc, Autodesk.Revit.DB.FamilyManager fm)
        {
            var pNote = fm.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            var pDescr = fm.get_Parameter(new Guid("f194bf60-b880-4217-b793-1e0c30dda5e9")); // КП_О_Наименование
            var pMark = fm.get_Parameter(new Guid("2204049c-d557-4dfc-8d70-13f19715e46d")); // КП_О_Марка
            var pManuf = fm.get_Parameter(new Guid("a8cdbf7b-d60a-485e-a520-447d2055f351")); // КП_О_Завод-изготовитель

            string note = FirstNonEmptyByTypes(doc, fm, pNote) ?? "";
            string descr = FirstNonEmptyByTypes(doc, fm, pDescr) ?? "";
            string mark = FirstNonEmptyByTypes(doc, fm, pMark) ?? "";
            string manuf = FirstNonEmptyByTypes(doc, fm, pManuf) ?? "";

            return "{"
                 + "\"Описание\":\"" + JsonEscape(descr) + "\","
                 + "\"Марка\":\"" + JsonEscape(mark) + "\","
                 + "\"Производитель\":\"" + JsonEscape(manuf) + "\","
                 + "\"Примечание\":\"" + JsonEscape(note) + "\""
                 + "}";
        }




        // IMPORT_INFO. Первый непустой текст среди всех типов семейства для указанного FamilyParameter
        private static string FirstNonEmptyByTypes(Document doc, Autodesk.Revit.DB.FamilyManager fm, FamilyParameter param)
        {
            if (fm == null || param == null || fm.Types == null) return null;

            foreach (FamilyType t in fm.Types)
            {
                if (t == null) continue;
                var val = GetFamilyParamStringValue(t, param);
                if (!string.IsNullOrWhiteSpace(val))
                    return val;
            }

            bool isInstance = IsInstanceParamCompat(param);

            string formula = GetFormulaCompat(fm, param);
            bool hasFormula = !string.IsNullOrWhiteSpace(formula);

            if (isInstance && hasFormula)
            {
                return null;
            }
            if (isInstance && !hasFormula)
            {
                return null;
            }
            if (!isInstance && hasFormula)
            {
                return null;
            }

            return null;
        }

        // IMPORT_INFO. Определение параметра экземпляра
        private static bool IsInstanceParamCompat(FamilyParameter fp)
        {
            if (fp == null) return false;
            var prop = typeof(FamilyParameter).GetProperty("IsInstance");
            if (prop != null && prop.CanRead)
            {
                try { return (bool)prop.GetValue(fp); } catch { }
            }
            return false; 
        }

        // IMPORT_INFO. Определение наличия формулы
        private static string GetFormulaCompat(Autodesk.Revit.DB.FamilyManager fm, FamilyParameter fp)
        {
            if (fm == null || fp == null) return null;

            var prop = typeof(FamilyParameter).GetProperty("Formula");
            if (prop != null && prop.CanRead)
            {
                try { return prop.GetValue(fp) as string; } catch { }
            }

            var mi = typeof(Autodesk.Revit.DB.FamilyManager).GetMethod("GetFormula", new[] { typeof(FamilyParameter) });
            if (mi != null)
            {
                try { return mi.Invoke(fm, new object[] { fp }) as string; } catch { }
            }
            return null;
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
                _records = LoadFamilyManagerRecords(DB_PATH);
                RebuildSearchIndex();
                RefreshScenario();                             
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
                        bool doStage = optDlg.DoStage;
                        bool doProject = optDlg.DoProject;
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
                                    sb.AppendLine($"Обновлено «IMPORT_INFO»: {handler.Updated} из {handler.Selected}.");
                                    sb.AppendLine($"Ошибок: {handler.Errors}.\n");
                                    sb.AppendLine("Для продолжения нажмите «Закрыть», а потом «тыкните» в любом месте интерфейса Revit, чтобы вернуться в контекст Revit uiApp. Простите за данные неудобства :(");

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
                            if (doDepartment || doStage || doProject || doImage) 
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
    }
}