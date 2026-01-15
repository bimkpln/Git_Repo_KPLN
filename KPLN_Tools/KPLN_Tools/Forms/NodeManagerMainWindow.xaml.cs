using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Xml.Linq;
using Transform = Autodesk.Revit.DB.Transform;


namespace KPLN_Tools.Forms
{
    // СТРУКТУРА КАТЕГОРИИ УЗЛОВ
    public class NodeCategoryUi : INotifyPropertyChanged
    {
        public string Id { get; set; }
        public string Title { get; set; }
        public int CatId { get; set; }

        public ObservableCollection<NodeCategoryUi> Children { get; } = new ObservableCollection<NodeCategoryUi>();

        public NodeCategoryUi Parent { get; set; }

        public int Depth
        {
            get
            {
                int d = 1;
                var p = Parent;
                while (p != null) { d++; p = p.Parent; }
                return d;
            }
        }

        public bool IsNonCatRoot => Id == "0";

        private int _directCount;
        public int DirectCount
        {
            get => _directCount;
            set
            {
                if (_directCount != value)
                {
                    _directCount = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(DisplayTitle));
                }
            }
        }

        private int _totalCount;
        public int TotalCount
        {
            get => _totalCount;
            set
            {
                if (_totalCount != value)
                {
                    _totalCount = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(DisplayTitle));
                }
            }
        }

        private bool _isExpanded;
        public bool IsExpanded
        {
            get => _isExpanded;
            set
            {
                if (_isExpanded != value)
                {
                    _isExpanded = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged();
                }
            }
        }

        public string DisplayTitle => $"{Title} ({TotalCount})";
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propName = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propName));
    }

    // СТРУКТУРА УЗЛА
    public class NodeElementUi
    {
        public long Id { get; set; }
        public string Name { get; set; }
        public ImageSource Image { get; set; }

        public string ScaleText { get; set; }
        public double? ScaleValue { get; set; }

        public int? CatId { get; set; }
        public string SubcatId { get; set; }
        public string CategoryPath { get; set; }

        private string _tags;
        public string Tags
        {
            get => _tags;
            set => _tags = value;
        }

        // Нормализованная строка тегов: DWG первым, алфавит, без дублей
        public string TagsNormalized
        {
            get => TagHelper.NormalizeTagsString(_tags);
            set
            {
                // Если вдруг кто-то привяжется TwoWay к TagsNormalized —
                // просто кладём значение в Tags.
                _tags = value;
            }
        }

        public string Comment { get; set; }

        public string TimeCreate { get; set; }
        public string TimeUpdate { get; set; }
        public string UserCreate { get; set; }
        public string UserUpdate { get; set; }

        public string RvtFileName { get; set; }

        public ObservableCollection<KeyValuePair<string, string>> PropParameters { get; set; }
            = new ObservableCollection<KeyValuePair<string, string>>();
    }

    // КОНВЕКТОРЫ ТЕГОВ
    public class TagsToListConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var s = value as string;
            if (string.IsNullOrWhiteSpace(s))
                return Array.Empty<string>();

            return TagHelper.SplitAndSortTags(s);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class BoolToVisibilityInverseConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool flag = value is bool b && b;
            return flag ? System.Windows.Visibility.Collapsed : System.Windows.Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    // ТЕГИ
    public static class TagHelper
    {
        public static List<string> SplitAndSortTags(string tags)
        {
            if (string.IsNullOrWhiteSpace(tags))
                return new List<string>();

            var list = tags
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(t => t.Trim())
                .Where(t => !string.IsNullOrWhiteSpace(t))
                .Distinct(StringComparer.InvariantCultureIgnoreCase)
                .ToList();

            return list
                .OrderBy(t => t.Equals("DWG", StringComparison.InvariantCultureIgnoreCase) ? 0 : 1)
                .ThenBy(t => t, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        public static string NormalizeTagsString(string tags)
        {
            var list = SplitAndSortTags(tags);
            return list.Count == 0 ? string.Empty : string.Join(", ", list);
        }
    }


    // ОСНОВНОЕ ОКНО
    public partial class MainWindowNodeManager : Window
    {
        private const string DbPath = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_NodeManager.db";
        public bool IsSuperUser { get; }
        private readonly string _userName;
        private readonly int _subDB;

        private readonly UIApplication _uiapp;
        private readonly UIDocument _uidoc;
        private readonly Document _doc;

        private readonly CopyNodeToViewHandler _copyHandler = new CopyNodeToViewHandler();
        private readonly ExternalEvent _copyEvent;

        private readonly PlaceDraftingViewOnSheetHandler _placeViewHandler;
        private readonly ExternalEvent _placeViewEvent;

        private static string MakeKey(int catId, string subId)
        {
            return $"{catId}|{subId ?? "0"}";
        }

        private class NodeCategoryRow
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public string SubcatJson { get; set; }
        }

        private class NodeElementRow
        {
            public int CatId { get; set; }
            public string SubcatId { get; set; }
        }

        private class SubcatDto
        {
            [JsonProperty("ID")]
            public string Id { get; set; }

            [JsonProperty("EL")]
            public string Name { get; set; }

            [JsonProperty("CH")]
            public List<SubcatDto> Children { get; set; }
        }

        public ObservableCollection<NodeCategoryUi> CategoryTree { get; } = new ObservableCollection<NodeCategoryUi>();
        public ObservableCollection<NodeElementUi> Elements { get; } = new ObservableCollection<NodeElementUi>();

        private NodeElementUi _currentElement;
        private readonly Dictionary<string, NodeCategoryUi> _nodeIndex = new Dictionary<string, NodeCategoryUi>();

        public ICollectionView ElementsView { get; }
        private NodeCategoryUi _currentSelectedNode;
        private bool _isSearchActive;

        public bool IsEditTagsMode
        {
            get => (bool)GetValue(IsEditTagsModeProperty);
            set => SetValue(IsEditTagsModeProperty, value);
        }

        public static readonly DependencyProperty IsEditTagsModeProperty =
            DependencyProperty.Register(
                nameof(IsEditTagsMode),
                typeof(bool),
                typeof(MainWindowNodeManager),
                new PropertyMetadata(false));

        public MainWindowNodeManager(UIApplication uiapp, UIDocument uidoc)
        {
            InitializeComponent();

            IsSuperUser = IsUserInTestList();
            BtnSettings.IsEnabled = IsSuperUser;

            var currentUser = KPLN_Library_SQLiteWorker.DBMainService.CurrentDBUser;
            if (currentUser != null)
            {
                _userName = GetUserNameFromMainDb(currentUser.Id);
                _subDB = currentUser.SubDepartmentId;
            }

            BtnCopyInView.IsEnabled = false;
            BtnCopyElements.IsEnabled = false;
            TxtComment.IsEnabled = false;
            BtnSaveComment.IsEnabled = false;
            TxtTags.IsEnabled = false;
            BtnAddTag.IsEnabled = false;
            BtnChangeCategory.IsEnabled = false;
            BtnReplacePreview.IsEnabled = false;

            _uiapp = uiapp;
            _uidoc = uidoc;
            _doc = uidoc.Document;

            _copyEvent = ExternalEvent.Create(_copyHandler);
            _placeViewHandler = new PlaceDraftingViewOnSheetHandler();
            _placeViewEvent = ExternalEvent.Create(_placeViewHandler);

            ElementsView = CollectionViewSource.GetDefaultView(Elements);
            DataContext = this;

            LoadCategoriesTree();
            LoadElementCounts();

            if (_currentSelectedNode != null)
            {
                LoadElementsForNode(_currentSelectedNode);
            }

            Loaded += MainWindowNodeManager_Loaded;
        }

        private void MainWindowNodeManager_Loaded(object sender, RoutedEventArgs e)
        {
            if (_currentSelectedNode != null)
            {
                _currentSelectedNode.IsExpanded = true;
                _currentSelectedNode.IsSelected = true;

                LoadElementsForNode(_currentSelectedNode);
            }
        }


        private bool IsUserInTestList()
        {
#if !Revit2020 && !Debug2020
            return false;
#else
            try
            {
                var currentUser = KPLN_Library_SQLiteWorker.DBMainService.CurrentDBUser;
                if (currentUser == null)
                    return false;

                var allowedIds = GetAllowedUserIdsFromDb();
                if (allowedIds == null || allowedIds.Count == 0)
                    return false;

                return allowedIds.Contains(currentUser.Id);
            }
            catch
            {
                return false;
            }
#endif
        }


        private List<int> GetAllowedUserIdsFromDb()
        {
            var result = new List<int>();

            try
            {
                using (var conn = new SQLiteConnection("Data Source=" + DbPath + ";Version=3;FailIfMissing=False;"))
                {
                    conn.Open();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "SELECT LIST FROM userAccess LIMIT 1;";
                        var obj = cmd.ExecuteScalar();
                        if (obj == null || obj == DBNull.Value)
                            return result;

                        var listStr = Convert.ToString(obj);
                        if (string.IsNullOrWhiteSpace(listStr))
                            return result;

                        var parts = listStr
                            .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(p => p.Trim());

                        foreach (var part in parts)
                        {
                            if (int.TryParse(part, out int id))
                            {
                                result.Add(id);
                            }
                        }
                    }
                    conn.Close();
                }
            }
            catch
            {
            }

            return result.Distinct().ToList();
        }

        private string GetUserNameFromMainDb(int userId)
        {
            string path = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_Loader_MainDB.db";

            const int maxAttempts = 3;
            const int delayMs = 3000;

            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    using (var conn = new SQLiteConnection($"Data Source={path};Version=3;FailIfMissing=False;"))
                    {
                        conn.Open();

                        using (var cmd = conn.CreateCommand())
                        {
                            cmd.CommandText = @"
                        SELECT Surname, Name
                        FROM Users
                        WHERE ID = @id
                        LIMIT 1;";

                            cmd.Parameters.AddWithValue("@id", userId);

                            using (var reader = cmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    string surname = reader["Surname"]?.ToString() ?? "";
                                    string name = reader["Name"]?.ToString() ?? "";

                                    return $"{surname} {name}".Trim();
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (attempt == maxAttempts)
                    {
                        TaskDialog.Show("Ошибка чтения Users",
                            $"Не удалось получить имя пользователя после {maxAttempts} попыток.\nОшибка:\n{ex.Message}");
                    }
                }

                if (attempt < maxAttempts)
                {
                    System.Threading.Thread.Sleep(delayMs);
                }
            }

            return null;
        }






        private List<string> LoadRvtPathsFromDb()
        {
            var result = new List<string>();

            try
            {
                using (var conn = new SQLiteConnection("Data Source=" + DbPath + ";Version=3;FailIfMissing=False;"))
                {
                    conn.Open();
                    using (var cmd = conn.CreateCommand())
                    {
                        if (_subDB == 8)
                        {
                            cmd.CommandText = @"
                                SELECT PATH
                                FROM RVTPath
                                WHERE PATH IS NOT NULL AND TRIM(PATH) <> '';";
                        }
                        else
                        {
                            cmd.CommandText = @"
                                SELECT PATH
                                FROM RVTPath
                                WHERE SUBDP = @subdp
                                  AND PATH IS NOT NULL AND TRIM(PATH) <> '';";
                            cmd.Parameters.AddWithValue("@subdp", _subDB);
                        }

                        using (var r = cmd.ExecuteReader())
                        {
                            while (r.Read())
                            {
                                if (r.IsDBNull(0))
                                    continue;

                                var path = r.GetString(0);
                                if (string.IsNullOrWhiteSpace(path))
                                    continue;

                                path = path.Trim();

                                try
                                {
                                    path = System.IO.Path.GetFullPath(path);
                                }
                                catch
                                {

                                }

                                result.Add(path);
                            }
                        }
                    }
                    conn.Close();
                }
            }
            catch
            {
            }

            return result
                .Distinct(StringComparer.InvariantCultureIgnoreCase)
                .ToList();
        }



        //////////////////////////////// КАТЕГОРИИ УЗЛОВ
        private List<NodeCategoryRow> LoadNodeCategories()
        {
            var result = new List<NodeCategoryRow>();

            using (var conn = new SQLiteConnection("Data Source=" + DbPath + ";Version=3;FailIfMissing=False;"))
            {
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT ID, NAME, SUBCAT_JSON FROM nodeCategory ORDER BY ID;";
                    using (var r = cmd.ExecuteReader())
                    {
                        while (r.Read())
                        {
                            var row = new NodeCategoryRow
                            {
                                Id = r.GetInt32(0),
                                Name = r.IsDBNull(1) ? string.Empty : r.GetString(1),
                                SubcatJson = r.IsDBNull(2) ? null : r.GetString(2)
                            };
                            result.Add(row);
                        }
                    }
                }
                conn.Close();
            }

            return result;
        }

        private void LoadCategoriesTree()
        {
            CategoryTree.Clear();
            _nodeIndex.Clear();

            var categories = LoadNodeCategories();

            NodeCategoryUi globalNoCat = null;

            // "Без категории" — только суперюзеру
            if (IsSuperUser)
            {
                globalNoCat = new NodeCategoryUi
                {
                    CatId = 0,
                    Id = "0",
                    Title = "Без категории"
                };
                CategoryTree.Add(globalNoCat);
                _nodeIndex[MakeKey(0, "0")] = globalNoCat;
            }

            _currentSelectedNode = null;

            // ===== BIM =====
            if (_subDB == 8)
            {
                // BIM видит все категории, как "старое" поведение
                foreach (var cat in categories)
                {
                    var rootNode = new NodeCategoryUi
                    {
                        CatId = cat.Id,
                        Id = cat.Id.ToString(CultureInfo.InvariantCulture),
                        Title = cat.Name
                    };

                    CategoryTree.Add(rootNode);
                    _nodeIndex[MakeKey(cat.Id, "0")] = rootNode;

                    if (!string.IsNullOrWhiteSpace(cat.SubcatJson))
                    {
                        try
                        {
                            var subList = JsonConvert.DeserializeObject<List<SubcatDto>>(cat.SubcatJson);
                            if (subList != null)
                            {
                                foreach (var s in subList)
                                {
                                    AttachSubtree(cat.Id, rootNode, s);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            TaskDialog.Show("KPLN. Менеджер узлов",
                                $"Ошибка при разборе SUBCAT_JSON для категории '{cat.Name}':\n{ex.Message}");
                        }
                    }
                }

                // По умолчанию выбираем либо BIM-категорию, либо первую
                _currentSelectedNode =
                    CategoryTree.FirstOrDefault(n => n.CatId == 8)
                    ?? CategoryTree.FirstOrDefault();
            }
            else
            {
                // ===== Обычные отделы (2–7) =====
                var deptRow = categories.FirstOrDefault(c => c.Id == _subDB);

                if (deptRow != null)
                {
                    var rootNode = new NodeCategoryUi
                    {
                        CatId = deptRow.Id,
                        Id = deptRow.Id.ToString(CultureInfo.InvariantCulture),
                        Title = deptRow.Name
                    };

                    CategoryTree.Add(rootNode);
                    _nodeIndex[MakeKey(deptRow.Id, "0")] = rootNode;

                    if (!string.IsNullOrWhiteSpace(deptRow.SubcatJson))
                    {
                        try
                        {
                            var subList = JsonConvert.DeserializeObject<List<SubcatDto>>(deptRow.SubcatJson);
                            if (subList != null)
                            {
                                foreach (var s in subList)
                                {
                                    AttachSubtree(deptRow.Id, rootNode, s);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            TaskDialog.Show("KPLN. Менеджер узлов",
                                $"Ошибка при разборе SUBCAT_JSON для категории '{deptRow.Name}':\n{ex.Message}");
                        }
                    }

                    _currentSelectedNode = rootNode;
                }
                else
                {
                    foreach (var cat in categories)
                    {
                        var rootNode = new NodeCategoryUi
                        {
                            CatId = cat.Id,
                            Id = cat.Id.ToString(CultureInfo.InvariantCulture),
                            Title = cat.Name
                        };

                        CategoryTree.Add(rootNode);
                        _nodeIndex[MakeKey(cat.Id, "0")] = rootNode;

                        if (!string.IsNullOrWhiteSpace(cat.SubcatJson))
                        {
                            try
                            {
                                var subList = JsonConvert.DeserializeObject<List<SubcatDto>>(cat.SubcatJson);
                                if (subList != null)
                                {
                                    foreach (var s in subList)
                                    {
                                        AttachSubtree(cat.Id, rootNode, s);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                TaskDialog.Show("KPLN. Менеджер узлов",
                                    $"Ошибка при разборе SUBCAT_JSON для категории '{cat.Name}':\n{ex.Message}");
                            }
                        }
                    }

                    _currentSelectedNode = CategoryTree.FirstOrDefault();
                }
            }
        }












        private void AttachSubtree(int catId, NodeCategoryUi parent, SubcatDto dto)
        {
            var node = new NodeCategoryUi
            {
                CatId = catId,
                Id = dto.Id,
                Title = dto.Name,
                Parent = parent
            };

            parent.Children.Add(node);
            _nodeIndex[MakeKey(catId, dto.Id)] = node;

            if (dto.Children != null)
            {
                foreach (var child in dto.Children)
                {
                    AttachSubtree(catId, node, child);
                }
            }
        }






        private List<NodeElementRow> LoadNodeElements()
        {
            var result = new List<NodeElementRow>();

            using (var conn = new SQLiteConnection("Data Source=" + DbPath + ";Version=3;FailIfMissing=False;"))
            {
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT CAT, SUBCAT FROM nodeManager;";

                    using (var r = cmd.ExecuteReader())
                    {
                        while (r.Read())
                        {
                            int cat = r.IsDBNull(0) ? 0 : r.GetInt32(0);
                            string subcat = r.IsDBNull(1) ? null : r.GetString(1);

                            if (_subDB != 8)
                            {
                                if (cat != _subDB && !(IsSuperUser && cat == 0))
                                    continue;
                            }

                            result.Add(new NodeElementRow
                            {
                                CatId = cat,
                                SubcatId = subcat
                            });
                        }
                    }
                }
                conn.Close();
            }

            return result;
        }









        private List<string> GetSubcatIdsRecursive(NodeCategoryUi node)
        {
            var list = new List<string>();
            CollectSubcatIds(node, list);
            return list;
        }

        private void CollectSubcatIds(NodeCategoryUi node, List<string> list)
        {
            if (!string.IsNullOrEmpty(node.Id))
                list.Add(node.Id);

            foreach (var child in node.Children)
                CollectSubcatIds(child, list);
        }









        private void LoadElementCounts()
        {
            foreach (var node in _nodeIndex.Values)
            {
                node.DirectCount = 0;
                node.TotalCount = 0;
            }

            var elements = LoadNodeElements();

            foreach (var el in elements)
            {
                int catId = el.CatId;
                string subId = string.IsNullOrWhiteSpace(el.SubcatId)
                    ? "0"
                    : el.SubcatId.Trim();

                NodeCategoryUi node = null;

                if (!_nodeIndex.TryGetValue(MakeKey(catId, subId), out node))
                {
                    _nodeIndex.TryGetValue(MakeKey(catId, "0"), out node);
                }

                if (node != null)
                {
                    node.DirectCount++;
                }
            }

            foreach (var root in CategoryTree)
            {
                RecalculateTotals(root);
            }
        }






        private int RecalculateTotals(NodeCategoryUi node)
        {
            int sum = node.DirectCount;

            foreach (var child in node.Children)
            {
                sum += RecalculateTotals(child);
            }

            node.TotalCount = sum;
            return sum;
        }

        private void CategoriesTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (_isSearchActive)
                return;

            _currentSelectedNode = e.NewValue as NodeCategoryUi;

            if (_currentSelectedNode != null)
            {
                LoadElementsForNode(_currentSelectedNode);
            }
        }

        private static T VisualUpwardSearch<T>(DependencyObject source) where T : DependencyObject
        {
            while (source != null && !(source is T))
            {
                source = VisualTreeHelper.GetParent(source);
            }
            return source as T;
        }

        private void CategoriesTree_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (!_isSearchActive)
                return;

            var clickedItem = VisualUpwardSearch<TreeViewItem>(e.OriginalSource as DependencyObject);
            if (clickedItem == null)
                return;

            var node = clickedItem.DataContext as NodeCategoryUi;
            if (node == null)
                return;

            _isSearchActive = false;
            SearchTextBox.Text = string.Empty;

            _currentSelectedNode = node;
            LoadElementsForNode(node);
        }

        //////////////////////////////// УЗЛЫ
        private void LoadElementsForNode(NodeCategoryUi node)
        {
            Elements.Clear();

            if (node == null)
            {
                NoElementsText.Visibility = System.Windows.Visibility.Visible;
                ElementsView.Refresh();
                return;
            }

            using (var conn = new SQLiteConnection("Data Source=" + DbPath + ";Version=3;FailIfMissing=False;"))
            {
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    if (node.CatId == 0 && node.Parent == null)
                    {
                        cmd.CommandText = @"
                            SELECT ID, NAME, PIC, PROP, CAT, SUBCAT, TAGS, COMMENT, TIME_CREATE, TIME_UPDATE, USER_CREATE, USER_UPDATE, RVT_PATH
                            FROM nodeManager
                            WHERE CAT IS NULL OR CAT = 0;";
                    }
                    else if (node.Parent == null)
                    {
                        cmd.CommandText = @"
                            SELECT ID, NAME, PIC, PROP, CAT, SUBCAT, TAGS, COMMENT, TIME_CREATE, TIME_UPDATE, USER_CREATE, USER_UPDATE, RVT_PATH
                            FROM nodeManager
                            WHERE CAT = @cat;";

                        cmd.Parameters.AddWithValue("@cat", node.CatId);
                    }
                    else
                    {
                        int catId = node.CatId;
                        var subIds = GetSubcatIdsRecursive(node);

                        if (subIds.Count == 0)
                        {
                            NoElementsText.Visibility = System.Windows.Visibility.Visible;
                            ElementsView.Refresh();
                            return;
                        }

                        var paramNames = new List<string>();
                        for (int i = 0; i < subIds.Count; i++)
                        {
                            string pName = "@s" + i;
                            paramNames.Add(pName);
                        }

                        cmd.CommandText = $@"
                            SELECT ID, NAME, PIC, PROP, CAT, SUBCAT, TAGS, COMMENT, TIME_CREATE, TIME_UPDATE, USER_CREATE, USER_UPDATE, RVT_PATH
                            FROM nodeManager
                            WHERE CAT = @cat
                              AND SUBCAT IN ({string.Join(", ", paramNames)});";

                        cmd.Parameters.AddWithValue("@cat", catId);
                        for (int i = 0; i < subIds.Count; i++)
                        {
                            cmd.Parameters.AddWithValue("@s" + i, subIds[i]);
                        }
                    }

                    using (var r = cmd.ExecuteReader())
                    {
                        while (r.Read())
                        {
                            long id = r.GetInt64(0);
                            string name = r.IsDBNull(1) ? string.Empty : r.GetString(1);
                            byte[] picBytes = null;
                            if (!r.IsDBNull(2))
                                picBytes = (byte[])r[2];
                            string propJson = r.IsDBNull(3) ? null : r.GetString(3);
                            int? catId = r.IsDBNull(4) ? (int?)null : r.GetInt32(4);
                            string subcatId = r.IsDBNull(5) ? null : r.GetString(5);
                            string tags = r.IsDBNull(6) ? null : r.GetString(6);
                            string comment = r.IsDBNull(7) ? null : r.GetString(7);

                            string timeCr = r.IsDBNull(8) ? null : r.GetString(8);
                            string timeUpd = r.IsDBNull(9) ? null : r.GetString(9);
                            string userCr = r.IsDBNull(10) ? null : r.GetString(10);
                            string userUpd = r.IsDBNull(11) ? null : r.GetString(11);

                            string rvtPath = r.IsDBNull(12) ? null : r.GetString(12);

                            string fileName = null;
                            if (!string.IsNullOrWhiteSpace(rvtPath))
                            {
                                try
                                {
                                    fileName = System.IO.Path.GetFileName(rvtPath);
                                }
                                catch { }
                            }

                            string scaleText = null;
                            double? scaleValue = null;

                            var propParams = new ObservableCollection<KeyValuePair<string, string>>();

                            if (!string.IsNullOrWhiteSpace(propJson))
                            {
                                try
                                {
                                    var jo = JObject.Parse(propJson);

                                    foreach (var p in jo.Properties())
                                    {
                                        propParams.Add(new KeyValuePair<string, string>(
                                            p.Name,
                                            p.Value?.ToString()
                                        ));
                                    }

                                    var scale = jo["МАСШТАБ"]?.ToString();
                                    if (!string.IsNullOrWhiteSpace(scale))
                                    {
                                        scaleText = scale;

                                        var parts = scale.Split(':');
                                        if (parts.Length == 2 &&
                                            double.TryParse(parts[1], NumberStyles.Any, CultureInfo.InvariantCulture, out double denom))
                                        {
                                            scaleValue = denom;
                                        }
                                    }
                                }
                                catch { }
                            }

                            string categoryPath = GetCategoryPath(catId, subcatId);

                            var element = new NodeElementUi
                            {
                                Id = id,
                                Name = name,
                                Image = LoadImageFromBytes(picBytes),
                                ScaleText = scaleText,
                                ScaleValue = scaleValue,
                                CatId = catId,
                                SubcatId = subcatId,
                                Tags = tags,
                                Comment = comment,
                                TimeCreate = timeCr,
                                TimeUpdate = timeUpd,
                                UserCreate = userCr,
                                UserUpdate = userUpd,
                                CategoryPath = categoryPath,
                                RvtFileName = fileName,
                                PropParameters = propParams
                            };

                            Elements.Add(element);
                        }
                    }
                }
                conn.Close();
            }

            NoElementsText.Visibility = Elements.Count == 0
            ? System.Windows.Visibility.Visible
            : System.Windows.Visibility.Collapsed;

            ApplySorting();
        }


        //////////////////////////////// ДЕЙСТВИЯ + РЕЗУЛЬТАТЫ ВЫВОДА
        private void BtnSearch_Click(object sender, RoutedEventArgs e)
        {
            string term = SearchTextBox.Text?.Trim();

            if (string.IsNullOrEmpty(term))
            {
                _isSearchActive = false;
                LoadElementsForNode(_currentSelectedNode);
            }
            else
            {
                _isSearchActive = true;
                SearchElementsByName(term);
            }
        }

        private void SearchElementsByName(string term)
        {
            Elements.Clear();

            using (var conn = new SQLiteConnection("Data Source=" + DbPath + ";Version=3;FailIfMissing=False;"))
            {
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
                        SELECT ID, NAME, PIC, PROP, CAT, SUBCAT, TAGS, COMMENT, TIME_CREATE, TIME_UPDATE, USER_CREATE, USER_UPDATE, RVT_PATH
                        FROM nodeManager;";

                    using (var r = cmd.ExecuteReader())
                    {
                        while (r.Read())
                        {
                            long id = r.GetInt64(0);
                            string name = r.IsDBNull(1) ? string.Empty : r.GetString(1);

                            if (string.IsNullOrEmpty(name) ||
                                name.IndexOf(term, StringComparison.OrdinalIgnoreCase) < 0)
                            {
                                continue;
                            }

                            byte[] picBytes = null;
                            if (!r.IsDBNull(2))
                                picBytes = (byte[])r[2];

                            string propJson = r.IsDBNull(3) ? null : r.GetString(3);
                            int? catId = r.IsDBNull(4) ? (int?)null : r.GetInt32(4);
                            string subcatId = r.IsDBNull(5) ? null : r.GetString(5);
                            string tags = r.IsDBNull(6) ? null : r.GetString(6);
                            string comment = r.IsDBNull(7) ? null : r.GetString(7);

                            string timeCr = r.IsDBNull(8) ? null : r.GetString(8);
                            string timeUpd = r.IsDBNull(9) ? null : r.GetString(9);
                            string userCr = r.IsDBNull(10) ? null : r.GetString(10);
                            string userUpd = r.IsDBNull(11) ? null : r.GetString(11);

                            string rvtPath = r.IsDBNull(12) ? null : r.GetString(12);

                            string fileName = null;
                            if (!string.IsNullOrWhiteSpace(rvtPath))
                            {
                                try { fileName = System.IO.Path.GetFileName(rvtPath); } catch { }
                            }

                            string scaleText = null;
                            double? scaleValue = null;

                            var propParams = new ObservableCollection<KeyValuePair<string, string>>();

                            if (!string.IsNullOrWhiteSpace(propJson))
                            {
                                try
                                {
                                    var jo = JObject.Parse(propJson);

                                    foreach (var p in jo.Properties())
                                    {
                                        propParams.Add(new KeyValuePair<string, string>(
                                            p.Name,
                                            p.Value?.ToString()
                                        ));
                                    }

                                    var scale = jo["МАСШТАБ"]?.ToString();
                                    if (!string.IsNullOrWhiteSpace(scale))
                                    {
                                        scaleText = scale;

                                        var parts = scale.Split(':');
                                        if (parts.Length == 2 &&
                                            double.TryParse(parts[1], NumberStyles.Any, CultureInfo.InvariantCulture, out double denom))
                                        {
                                            scaleValue = denom;
                                        }
                                    }
                                }
                                catch
                                {
                                }
                            }

                            string categoryPath = GetCategoryPath(catId, subcatId);

                            var element = new NodeElementUi
                            {
                                Id = id,
                                Name = name,
                                Image = LoadImageFromBytes(picBytes),
                                ScaleText = scaleText,
                                ScaleValue = scaleValue,
                                CatId = catId,
                                SubcatId = subcatId,
                                Tags = tags,
                                Comment = comment,
                                TimeCreate = timeCr,
                                TimeUpdate = timeUpd,
                                UserCreate = userCr,
                                UserUpdate = userUpd,
                                CategoryPath = categoryPath,
                                RvtFileName = fileName,
                                PropParameters = propParams
                            };

                            Elements.Add(element);
                        }
                    }
                }
                conn.Close();
            }

            NoElementsText.Visibility = Elements.Count == 0
                ? System.Windows.Visibility.Visible
                : System.Windows.Visibility.Collapsed;

            ElementsView.Refresh();
        }

        private void BtnFilter_Click(object sender, RoutedEventArgs e)
        {
            List<string> allTags;
            try
            {
                allTags = GetAllTagsFromDb();
            }
            catch (Exception ex)
            {
                TaskDialog.Show("KPLN. Менеджер узлов", "Ошибка при чтении тегов из БД:\n" + ex.Message);
                return;
            }

            if (allTags == null || allTags.Count == 0)
            {
                TaskDialog.Show("KPLN. Менеджер узлов", "В БД нет тегов для фильтрации.");
                return;
            }

            var dlg = new FilterByTagsWindowNodeManager(allTags)
            {
                Owner = this
            };

            var result = dlg.ShowDialog();
            if (result == true)
            {
                var selectedTags = dlg.ResultTags;
                if (selectedTags != null && selectedTags.Count > 0)
                {
                    _isSearchActive = true;
                    SearchTextBox.Text = string.Empty;

                    SearchElementsByTags(selectedTags);
                }
            }
        }

        private List<string> GetAllTagsFromDb()
        {
            var result = new List<string>();

            using (var conn = new SQLiteConnection("Data Source=" + DbPath + ";Version=3;FailIfMissing=False;"))
            {
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT TAGS FROM nodeManager WHERE TAGS IS NOT NULL AND TRIM(TAGS) <> '';";

                    using (var r = cmd.ExecuteReader())
                    {
                        while (r.Read())
                        {
                            if (r.IsDBNull(0))
                                continue;

                            var tagsStr = r.GetString(0);
                            if (string.IsNullOrWhiteSpace(tagsStr))
                                continue;

                            var parts = tagsStr
                                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                .Select(t => t.Trim())
                                .Where(t => !string.IsNullOrWhiteSpace(t));

                            result.AddRange(parts);
                        }
                    }
                }
                conn.Close();
            }

            return result
                .Distinct(StringComparer.InvariantCultureIgnoreCase)
                .ToList();
        }

        private void SearchElementsByTags(List<string> tagsToFind)
        {
            Elements.Clear();

            if (tagsToFind == null || tagsToFind.Count == 0)
            {
                NoElementsText.Visibility = System.Windows.Visibility.Visible;
                ElementsView.Refresh();
                return;
            }


            var selectedTagsNorm = tagsToFind
                .Where(t => !string.IsNullOrWhiteSpace(t))
                .Select(t => t.Trim())
                .Distinct(StringComparer.InvariantCultureIgnoreCase)
                .ToList();

            using (var conn = new SQLiteConnection("Data Source=" + DbPath + ";Version=3;FailIfMissing=False;"))
            {
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
                        SELECT ID, NAME, PIC, PROP, CAT, SUBCAT, TAGS, COMMENT, TIME_CREATE, TIME_UPDATE, USER_CREATE, USER_UPDATE, RVT_PATH
                        FROM nodeManager
                        WHERE TAGS IS NOT NULL AND TRIM(TAGS) <> '';";

                    using (var r = cmd.ExecuteReader())
                    {
                        while (r.Read())
                        {
                            long id = r.GetInt64(0);
                            string name = r.IsDBNull(1) ? string.Empty : r.GetString(1);
                            byte[] picBytes = null;
                            if (!r.IsDBNull(2))
                                picBytes = (byte[])r[2];
                            string propJson = r.IsDBNull(3) ? null : r.GetString(3);
                            int? catId = r.IsDBNull(4) ? (int?)null : r.GetInt32(4);
                            string subcatId = r.IsDBNull(5) ? null : r.GetString(5);
                            string tagsStr = r.IsDBNull(6) ? null : r.GetString(6);
                            string comment = r.IsDBNull(7) ? null : r.GetString(7);

                            string timeCr = r.IsDBNull(8) ? null : r.GetString(8);
                            string timeUpd = r.IsDBNull(9) ? null : r.GetString(9);
                            string userCr = r.IsDBNull(10) ? null : r.GetString(10);
                            string userUpd = r.IsDBNull(11) ? null : r.GetString(11);

                            string rvtPath = r.IsDBNull(12) ? null : r.GetString(12);

                            string fileName = null;
                            if (!string.IsNullOrWhiteSpace(rvtPath))
                            {
                                try { fileName = System.IO.Path.GetFileName(rvtPath); } catch { }
                            }

                            if (string.IsNullOrWhiteSpace(tagsStr))
                                continue;

                            var rowTags = tagsStr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(t => t.Trim()).
                                Where(t => !string.IsNullOrWhiteSpace(t)).Distinct(StringComparer.InvariantCultureIgnoreCase).ToList();

                            if (rowTags.Count == 0)
                                continue;

                            var rowSet = new HashSet<string>(rowTags, StringComparer.InvariantCultureIgnoreCase);

                            bool containsAll = selectedTagsNorm.All(t => rowSet.Contains(t));
                            if (!containsAll)
                                continue;

                            string scaleText = null;
                            double? scaleValue = null;

                            var propParams = new ObservableCollection<KeyValuePair<string, string>>();

                            if (!string.IsNullOrWhiteSpace(propJson))
                            {
                                try
                                {
                                    var jo = JObject.Parse(propJson);

                                    foreach (var p in jo.Properties())
                                    {
                                        propParams.Add(new KeyValuePair<string, string>(
                                            p.Name,
                                            p.Value?.ToString()
                                        ));
                                    }

                                    var scale = jo["МАСШТАБ"]?.ToString();
                                    if (!string.IsNullOrWhiteSpace(scale))
                                    {
                                        scaleText = scale;

                                        var parts = scale.Split(':');
                                        if (parts.Length == 2 &&
                                            double.TryParse(parts[1], NumberStyles.Any, CultureInfo.InvariantCulture, out double denom))
                                        {
                                            scaleValue = denom;
                                        }
                                    }
                                }
                                catch
                                {
                                }
                            }

                            string categoryPath = GetCategoryPath(catId, subcatId);

                            var element = new NodeElementUi
                            {
                                Id = id,
                                Name = name,
                                Image = LoadImageFromBytes(picBytes),
                                ScaleText = scaleText,
                                ScaleValue = scaleValue,
                                CatId = catId,
                                SubcatId = subcatId,
                                Tags = tagsStr,
                                Comment = comment,
                                TimeCreate = timeCr,
                                TimeUpdate = timeUpd,
                                UserCreate = userCr,
                                UserUpdate = userUpd,
                                CategoryPath = categoryPath,
                                RvtFileName = fileName,
                                PropParameters = propParams
                            };

                            Elements.Add(element);
                        }
                    }
                }
                conn.Close();
            }

            NoElementsText.Visibility = Elements.Count == 0
                ? System.Windows.Visibility.Visible
                : System.Windows.Visibility.Collapsed;

            ElementsView.Refresh();
        }

        private void BtnClearSearch_Click(object sender, RoutedEventArgs e)
        {
            SearchTextBox.Text = string.Empty;
            _isSearchActive = false;

            LoadElementsForNode(_currentSelectedNode);
        }

        private void SortCombo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (SortCombo == null)
                return;

            if (SortCombo.SelectedIndex != 0 && ScaleSortCombo != null && ScaleSortCombo.SelectedIndex != 0)
            {
                ScaleSortCombo.SelectedIndex = 0;
            }

            ApplySorting();
        }

        private void ScaleSortCombo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (ScaleSortCombo == null)
                return;

            if (ScaleSortCombo.SelectedIndex != 0 && SortCombo != null && SortCombo.SelectedIndex != 0)
            {
                SortCombo.SelectedIndex = 0;
            }

            ApplySorting();
        }

        private void ApplySorting()
        {
            if (ElementsView == null)
                return;

            ElementsView.SortDescriptions.Clear();

            ComboBoxItem scaleItem = null;
            string scaleTag = null;

            if (ScaleSortCombo != null)
            {
                scaleItem = ScaleSortCombo.SelectedItem as ComboBoxItem;
                scaleTag = scaleItem?.Tag as string;
            }

            if (scaleTag == "bigFirst")
            {
                ElementsView.SortDescriptions.Add(new SortDescription(nameof(NodeElementUi.ScaleValue), ListSortDirection.Ascending));
            }
            else if (scaleTag == "smallFirst")
            {
                ElementsView.SortDescriptions.Add(new SortDescription(nameof(NodeElementUi.ScaleValue), ListSortDirection.Descending));
            }

            ComboBoxItem nameItem = null;
            string nameTag = null;

            if (SortCombo != null)
            {
                nameItem = SortCombo.SelectedItem as ComboBoxItem;
                nameTag = nameItem?.Tag as string;
            }

            if (nameTag == "asc")
            {
                ElementsView.SortDescriptions.Add(new SortDescription(nameof(NodeElementUi.Name), ListSortDirection.Ascending));
            }
            else if (nameTag == "desc")
            {
                ElementsView.SortDescriptions.Add(new SortDescription(nameof(NodeElementUi.Name), ListSortDirection.Descending));
            }

            ElementsView.Refresh();
        }

        private void ElementsList_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            _currentElement = ElementsList.SelectedItem as NodeElementUi;
            bool hasSelection = _currentElement != null;

            BtnCopyInView.IsEnabled = hasSelection;
            BtnCopyElements.IsEnabled = hasSelection;

            bool canEdit = hasSelection && IsSuperUser;
            TxtTags.IsEnabled = canEdit;
            BtnAddTag.IsEnabled = canEdit;
            BtnChangeCategory.IsEnabled = canEdit;
            BtnReplacePreview.IsEnabled = canEdit;

            TxtComment.IsEnabled = canEdit;
            BtnSaveComment.IsEnabled = canEdit;

            IsEditTagsMode = false;
        }

        //////////////////////////////// СВОЙСТВА УЗЛА
        private string GetCategoryPath(int? catId, string subcatId)
        {
            if (!catId.HasValue || catId.Value == 0)
                return "Без категории";

            string subNorm = string.IsNullOrEmpty(subcatId) ? "0" : subcatId;

            if (!_nodeIndex.TryGetValue(MakeKey(catId.Value, subNorm), out var node))
                return $"Категория {catId}, подкатегория {subcatId ?? "—"}";

            var stack = new Stack<string>();
            var cur = node;
            while (cur != null)
            {
                stack.Push(cur.Title);
                cur = cur.Parent;
            }

            return string.Join(" > ", stack);
        }
     
        private void BtnSaveComment_Click(object sender, RoutedEventArgs e)
        {
            if (!IsSuperUser)
                return;

            var element = ElementsList.SelectedItem as NodeElementUi;
            if (element == null)
                return;

            try
            {
                using (var conn = new SQLiteConnection("Data Source=" + DbPath + ";Version=3;FailIfMissing=False;"))
                {
                    conn.Open();
                    using (var cmd = conn.CreateCommand())
                    {
                        string now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                        string uName = _userName ?? string.Empty;

                        cmd.CommandText = @"
                            UPDATE nodeManager 
                            SET COMMENT = @comment,
                                TIME_UPDATE = @timeUpd,
                                USER_UPDATE = @userUpd
                            WHERE ID = @id;";

                        cmd.Parameters.AddWithValue("@comment", element.Comment ?? string.Empty);
                        cmd.Parameters.AddWithValue("@timeUpd", now);
                        cmd.Parameters.AddWithValue("@userUpd", uName);
                        cmd.Parameters.AddWithValue("@id", element.Id);

                        cmd.ExecuteNonQuery();
                    }
                    conn.Close();
                }

                ReloadCurrentElement(element.Id);
                TaskDialog.Show("KPLN. Менеджер узлов", "Комментарий успешно обновлён.");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("KPLN. Менеджер узлов", "Ошибка при обновлении комментария:\n" + ex.Message);
            }

            this.Topmost = true;
            this.Topmost = false;
        }

        private void BtnSaveTags_Click(object sender, RoutedEventArgs e)
        {
            if (!IsSuperUser)
                return;

            var element = ElementsList.SelectedItem as NodeElementUi;
            if (element == null)
                return;

            var dlg = new NodeManagerInputTagWindow
            {
                Owner = this
            };

            if (dlg.ShowDialog() != true)
                return;

            var newTag = dlg.TagText;
            if (string.IsNullOrWhiteSpace(newTag))
                return;

            var currentList = TagHelper.SplitAndSortTags(element.Tags);

            currentList.Add(newTag.Trim());

            element.Tags = TagHelper.NormalizeTagsString(string.Join(", ", currentList));

            SaveTagsToDb(element);
            ReloadCurrentElement(element.Id);
        }


        private void SaveTagsToDb(NodeElementUi element)
        {
            try
            {
                using (var conn = new SQLiteConnection("Data Source=" + DbPath + ";Version=3;FailIfMissing=False;"))
                {
                    conn.Open();
                    using (var cmd = conn.CreateCommand())
                    {
                        string now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                        string uName = _userName ?? string.Empty;

                        cmd.CommandText = @"
                    UPDATE nodeManager 
                    SET TAGS = @tags,
                        TIME_UPDATE = @timeUpd,
                        USER_UPDATE = @userUpd
                    WHERE ID = @id;";

                        cmd.Parameters.AddWithValue("@tags", element.Tags ?? string.Empty);
                        cmd.Parameters.AddWithValue("@timeUpd", now);
                        cmd.Parameters.AddWithValue("@userUpd", uName);
                        cmd.Parameters.AddWithValue("@id", element.Id);

                        cmd.ExecuteNonQuery();
                    }
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("KPLN. Менеджер узлов", "Ошибка при обновлении тегов:\n" + ex.Message);
            }
        }

        private void BtnAddTag_Click(object sender, RoutedEventArgs e)
        {
            if (!IsSuperUser)
                return;

            var element = ElementsList.SelectedItem as NodeElementUi;
            if (element == null)
                return;

            List<string> allTags;
            try
            {
                allTags = GetAllTagsFromDb(); 
            }
            catch (Exception ex)
            {
                TaskDialog.Show("KPLN. Менеджер узлов", "Ошибка при чтении тегов из БД:\n" + ex.Message);
                return;
            }

            if (allTags == null || allTags.Count == 0)
            {
                TaskDialog.Show("KPLN. Менеджер узлов", "В БД пока нет тегов для выбора.");
                return;
            }

            var currentTags = TagHelper.SplitAndSortTags(element.Tags);

            var dlg = new NodeManagerManageTagsWindow(allTags, currentTags)
            {
                Owner = this
            };

            if (dlg.ShowDialog() == true)
            {
                var resultTags = dlg.ResultTags ?? new List<string>();

                element.Tags = TagHelper.NormalizeTagsString(
                    string.Join(", ", resultTags));

                SaveTagsToDb(element); 
                ReloadCurrentElement(element.Id);
            }
        }



        private void BtnChangeCategory_Click(object sender, RoutedEventArgs e)
        {
            if (!IsSuperUser)
                return;

            if (_currentElement == null)
                return;

            try
            {
                var dlg = new ChangeCategoryWindowNodeManager(CategoryTree, _currentElement.CatId, _currentElement.SubcatId, _currentElement.Name);
                dlg.Owner = this;

                if (dlg.ShowDialog() == true)
                {
                    int? newCatId = dlg.SelectedCatId;
                    string newSubcatId = dlg.SelectedSubcatId;

                    if (newCatId.HasValue)
                    {
                        try
                        {
                            using (var conn = new SQLiteConnection("Data Source=" + DbPath + ";Version=3;FailIfMissing=False;"))
                            {
                                conn.Open();
                                using (var cmd = conn.CreateCommand())
                                {
                                    string now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                                    string uName = _userName ?? string.Empty;

                                    cmd.CommandText = @"
                                        UPDATE nodeManager 
                                        SET CAT = @cat,
                                            SUBCAT = @sub,
                                            TIME_UPDATE = @timeUpd,
                                            USER_UPDATE = @userUpd
                                        WHERE ID = @id;";

                                    cmd.Parameters.AddWithValue("@cat", newCatId.Value);
                                    cmd.Parameters.AddWithValue("@sub", (object)newSubcatId ?? DBNull.Value);
                                    cmd.Parameters.AddWithValue("@timeUpd", now);
                                    cmd.Parameters.AddWithValue("@userUpd", uName);
                                    cmd.Parameters.AddWithValue("@id", _currentElement.Id);

                                    cmd.ExecuteNonQuery();
                                }
                                conn.Close();
                            }
                        }
                        catch (Exception ex)
                        {
                            TaskDialog.Show("KPLN. Менеджер узлов", "Ошибка при обновлении категории:\n" + ex.Message);
                            return;
                        }

                        _currentElement.CatId = newCatId;
                        _currentElement.SubcatId = newSubcatId;
                        _currentElement.CategoryPath = GetCategoryPath(newCatId, newSubcatId);

                        LoadElementCounts();

                        if (_currentSelectedNode != null)
                        {
                            LoadElementsForNode(_currentSelectedNode);
                        }
                    }
                }
            }
            finally { }
        }


















        private void BtnReplacePreview_Click(object sender, RoutedEventArgs e)
        {
            if (_currentElement == null)
                return;

            View targetView = null;
            bool openedTargetViewHere = false;

            try
            {
                this.Hide();

                string rvtPath = null;
                string propJson = null;

                using (var conn = new SQLiteConnection("Data Source=" + DbPath + ";Version=3;FailIfMissing=False;"))
                {
                    conn.Open();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "SELECT RVT_PATH, PROP FROM nodeManager WHERE ID = @id;";
                        cmd.Parameters.AddWithValue("@id", _currentElement.Id);

                        using (var r = cmd.ExecuteReader())
                        {
                            if (r.Read())
                            {
                                if (!r.IsDBNull(0)) rvtPath = r.GetString(0);
                                if (!r.IsDBNull(1)) propJson = r.GetString(1);
                            }
                        }
                    }
                    conn.Close();
                }

                if (string.IsNullOrWhiteSpace(rvtPath))
                {
                    TaskDialog.Show("KPLN. Менеджер узлов",
                        "Для выбранного узла в базе не заполнен путь к модели (RVT_PATH).\nОбновите БД, чтобы заполнить это поле.");
                    return;
                }

                if (string.IsNullOrWhiteSpace(propJson))
                {
                    TaskDialog.Show("KPLN. Менеджер узлов",
                        "Для выбранного узла в базе отсутствует PROP.\nОбновите БД, чтобы заполнить это поле.");
                    return;
                }

                string viewName = null;
                try
                {
                    var jo = JObject.Parse(propJson);
                    viewName = jo["ИМЯ ВИДА"]?.ToString();
                }
                catch { }

                if (string.IsNullOrWhiteSpace(viewName))
                {
                    TaskDialog.Show("KPLN. Менеджер узлов",
                        "В PROP для выбранного узла не найден ключ \"ИМЯ ВИДА\".\nОбновите БД, чтобы корректно заполнить параметры вида.");
                    return;
                }

                // --------------------------------------------------------------------
                // 1) Нормализуем путь из БД (обычный file path). Для RevitServer/Cloud
                // --------------------------------------------------------------------
                string targetFull;
                try { targetFull = Path.GetFullPath(rvtPath); }
                catch { targetFull = rvtPath; }

                bool looksLikeFilePath = targetFull.Contains(":\\") || targetFull.StartsWith(@"\\");
                var app = _uiapp.Application;

                Document localDoc = null;
                Document centralDoc = null;

                // --------------------------------------------------------------------
                // 2) Ищем среди УЖЕ открытых документов:
                //    - если документ workshared: сравниваем его CENTRAL path с targetFull (из БД)
                //    - при совпадении: если текущий документ != central -> это local (предпочитаем его)
                // --------------------------------------------------------------------
                foreach (Document d in app.Documents)
                {
                    try
                    {
                        if (d == null) continue;

                        if (d.IsWorkshared)
                        {
                            var mp = d.GetWorksharingCentralModelPath();
                            if (mp == null) continue;

                            string centralPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(mp);

                            string centralFull;
                            try { centralFull = Path.GetFullPath(centralPath); }
                            catch { centralFull = centralPath; }

                            if (string.IsNullOrEmpty(centralFull)) continue;

                            if (!string.Equals(centralFull, targetFull, StringComparison.InvariantCultureIgnoreCase))
                                continue;

                            string docFull;
                            try { docFull = Path.GetFullPath(d.PathName); }
                            catch { docFull = d.PathName; }

                            if (!string.IsNullOrEmpty(docFull) &&
                                string.Equals(docFull, centralFull, StringComparison.InvariantCultureIgnoreCase))
                            {
                                centralDoc = d; 
                            }
                            else
                            {
                                localDoc = d; 
                            }
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(d.PathName)) continue;

                            string dFull;
                            try { dFull = Path.GetFullPath(d.PathName); }
                            catch { dFull = d.PathName; }

                            if (!string.IsNullOrEmpty(dFull) &&
                                string.Equals(dFull, targetFull, StringComparison.InvariantCultureIgnoreCase))
                            {
                                centralDoc = d;
                            }
                        }
                    }
                    catch { }
                }

                Document targetDoc = localDoc ?? centralDoc;

                // --------------------------------------------------------------------
                // 3) Если документ НЕ открыт - открываем.
                //    Если файл похож на обычный путь и это workshared central - пробуем создать local и открыть local.
                // --------------------------------------------------------------------
                if (targetDoc == null)
                {
                    if (looksLikeFilePath && !File.Exists(targetFull))
                    {
                        TaskDialog.Show("KPLN. Менеджер узлов",
                            "Файл модели, указанный в RVT_PATH, не найден:\n" + targetFull);
                        return;
                    }

                    bool opened = false;

                    if (looksLikeFilePath)
                    {
                        try
                        {
                            ModelPath centralMp = ModelPathUtils.ConvertUserVisiblePathToModelPath(targetFull);

                            string safeUser = (_userName ?? Environment.UserName ?? "user").Trim();
                            foreach (char c in Path.GetInvalidFileNameChars())
                                safeUser = safeUser.Replace(c, '_');

                            string localDir = Path.Combine(Path.GetTempPath(), "KPLN_NodeManager_Local");
                            Directory.CreateDirectory(localDir);

                            string baseName = Path.GetFileNameWithoutExtension(targetFull);
                            string localPath = Path.Combine(localDir, $"{baseName}_{safeUser}_{DateTime.Now:yyyyMMdd_HHmmss}.rvt");

                            ModelPath localMp = ModelPathUtils.ConvertUserVisiblePathToModelPath(localPath);

                            WorksharingUtils.CreateNewLocal(centralMp, localMp);

                            UIDocument uiDocLocal = _uiapp.OpenAndActivateDocument(localPath);
                            targetDoc = uiDocLocal.Document;
                            opened = (targetDoc != null);
                        }
                        catch
                        {
                        }
                    }

                    if (!opened)
                    {
                        try
                        {
                            UIDocument uiDoc = _uiapp.OpenAndActivateDocument(targetFull);
                            targetDoc = uiDoc.Document;
                        }
                        catch (Exception ex)
                        {
                            TaskDialog.Show("KPLN. Менеджер узлов", "Не удалось открыть файл модели:\n" + targetFull + "\n\n" + ex.Message);
                            return;
                        }
                    }
                }

                if (targetDoc == null)
                {
                    TaskDialog.Show("KPLN. Менеджер узлов", "Не удалось определить целевой документ.");
                    return;
                }

                // --------------------------------------------------------------------
                // 4) Активируем документ
                // --------------------------------------------------------------------
                UIDocument udoc = _uiapp.ActiveUIDocument;
                try
                {
                    if (udoc == null || udoc.Document != targetDoc)
                    {
                        string activatePath = targetDoc.PathName;
                        if (!string.IsNullOrEmpty(activatePath))
                            udoc = _uiapp.OpenAndActivateDocument(activatePath);
                    }
                }
                catch { }

                // --------------------------------------------------------------------
                // 5) Ищем вид в targetDoc
                // --------------------------------------------------------------------
                try
                {
                    targetView = new FilteredElementCollector(targetDoc)
                        .OfClass(typeof(View))
                        .Cast<View>()
                        .FirstOrDefault(v => !v.IsTemplate &&
                                             string.Equals(v.Name, viewName, StringComparison.InvariantCultureIgnoreCase));
                }
                catch { }

                if (targetView == null)
                {
                    TaskDialog.Show("KPLN. Менеджер узлов",
                        "В документе\n" + (targetDoc.PathName ?? targetDoc.Title) +
                        "\nне найден вид с именем:\n\"" + viewName + "\".\nВозможно, он был удалён или переименован.");
                    return;
                }

                // --------------------------------------------------------------------
                // 6) Проверяем, открыт ли вид (не переоткрывая централь!)
                // --------------------------------------------------------------------
                bool viewIsOpen = false;

                try
                {
                    if (udoc == null || udoc.Document != targetDoc)
                    {
                        // ещё раз пытаемся активировать именно targetDoc
                        if (!string.IsNullOrEmpty(targetDoc.PathName))
                            udoc = _uiapp.OpenAndActivateDocument(targetDoc.PathName);
                    }

                    if (udoc != null)
                    {
                        foreach (var uv in udoc.GetOpenUIViews())
                        {
                            if (uv.ViewId == targetView.Id)
                            {
                                viewIsOpen = true;
                                break;
                            }
                        }
                    }
                }
                catch
                {
                    viewIsOpen = false;
                }

                // --------------------------------------------------------------------
                // 7) Если вид не открыт - открываем (активируем)
                // --------------------------------------------------------------------
                if (!viewIsOpen)
                {
                    try
                    {
                        if (udoc != null)
                        {
                            udoc.ActiveView = targetView;
                            openedTargetViewHere = true;
                        }
                    }
                    catch { }
                }

                // --------------------------------------------------------------------
                // 8) Захват изображения и запись в БД
                // --------------------------------------------------------------------
                var captureWindow = new ScreenCaptureWindow(1000, 800);
                captureWindow.Owner = this;

                var dialogResult = captureWindow.ShowDialog();
                if (dialogResult != true || captureWindow.CapturedBytes == null || captureWindow.CapturedBytes.Length == 0)
                {
                    TaskDialog.Show("KPLN. Менеджер узлов", "Изображение не было захвачено. Повторите попытку.");
                    return;
                }

                byte[] previewBytes = captureWindow.CapturedBytes;

                try
                {
                    using (var conn = new SQLiteConnection("Data Source=" + DbPath + ";Version=3;FailIfMissing=False;"))
                    {
                        conn.Open();
                        using (var cmd = conn.CreateCommand())
                        {
                            string now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                            string uName = _userName ?? string.Empty;

                            cmd.CommandText = @"
                                UPDATE nodeManager 
                                SET PIC = @pic,
                                    TIME_UPDATE = @timeUpd,
                                    USER_UPDATE = @userUpd
                                WHERE ID = @id;";

                            var pPic = cmd.CreateParameter();
                            pPic.ParameterName = "@pic";
                            pPic.DbType = DbType.Binary;
                            pPic.Value = (object)previewBytes ?? DBNull.Value;
                            cmd.Parameters.Add(pPic);

                            cmd.Parameters.AddWithValue("@timeUpd", now);
                            cmd.Parameters.AddWithValue("@userUpd", uName);
                            cmd.Parameters.AddWithValue("@id", _currentElement.Id);

                            cmd.ExecuteNonQuery();
                        }
                        conn.Close();
                    }

                    TaskDialog.Show("KPLN. Менеджер узлов", "Миниатюра успешно сохранена в БД.");
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("KPLN. Менеджер узлов", "Ошибка при сохранении миниатюры в базу данных:\n" + ex.Message);
                    return;
                }

                if (_currentSelectedNode != null)
                {
                    var savedId = _currentElement.Id;
                    LoadElementsForNode(_currentSelectedNode);

                    var newSel = Elements.FirstOrDefault(el => el.Id == savedId);
                    if (newSel != null)
                    {
                        ElementsList.SelectedItem = newSel;
                        _currentElement = newSel;
                    }

                    ReloadCurrentElement(savedId);
                }
            }
            finally
            {
                try
                {
                    if (openedTargetViewHere && targetView != null)
                    {
                        var udoc2 = _uiapp.ActiveUIDocument;
                        if (udoc2 != null && udoc2.Document == targetView.Document)
                        {
                            var uv = udoc2.GetOpenUIViews()
                                          ?.FirstOrDefault(v => v.ViewId == targetView.Id);
                            uv?.Close();
                        }
                    }
                }
                catch { }

                this.Show();
                this.Topmost = true;              
                this.Topmost = false;
            }
        }







   



















        private void ReloadCurrentElement(long elementId)
        {
            if (_currentSelectedNode == null)
                return;
            LoadElementCounts();

            LoadElementsForNode(_currentSelectedNode);

            var newSel = Elements.FirstOrDefault(el => el.Id == elementId);
            if (newSel != null)
            {
                ElementsList.SelectedItem = newSel;
                _currentElement = newSel;
            }
        }

        private ImageSource LoadImageFromBytes(byte[] data)
        {
            if (data == null || data.Length == 0)
                return null;

            try
            {
                var bmp = new BitmapImage();
                using (var ms = new MemoryStream(data))
                {
                    bmp.BeginInit();
                    bmp.CacheOption = BitmapCacheOption.OnLoad;
                    bmp.StreamSource = ms;
                    bmp.EndInit();
                    bmp.Freeze();
                }
                return bmp;
            }
            catch
            {
                return null;
            }
        }

        ////////////////////// НАСТРОЙКИ
        private void BtnSettings_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new SettingsWindowNodeManager
            {
                Owner = this,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };

            if (dlg.ShowDialog() == true)
            {
                switch (dlg.SelectedAction)
                {
                    case SettingsAction.UpdateDb:
                        string run = Guid.NewGuid().ToString();

                        var rvtPaths = LoadRvtPathsFromDb();
                        if (rvtPaths == null || rvtPaths.Count == 0)
                        {
                            TaskDialog.Show("KPLN. Менеджер узлов", "В таблице RVTPath нет ни одного пути. Добавьте пути к моделям и повторите попытку.");
                            break;
                        }
                        foreach (var path in rvtPaths)
                        {
                            HandleAction_UpdateDb(_uiapp, path, run, _userName);
                        }

                        FinalizeUpdateDb(run, rvtPaths, _subDB == 8);
                        TaskDialog.Show("KPLN. Менеджер узлов", "Операция завершена");
                        break;

                    case SettingsAction.AddDelCategory:
                        var w = new SettingsWindowNodeManagerCategory(_subDB);
                        w.Owner = this;
                        w.ShowDialog();
                        break;
                }

                this.Topmost = true;
                this.Topmost = false;

                LoadCategoriesTree();
                LoadElementCounts();
            }
        }

        public static void HandleAction_UpdateDb(UIApplication uiapp, string rvtPath, string runToken, string userName)
        {
            Document doc = null;
            bool docOpenedHere = false;

            try
            {
                var app = uiapp.Application;
                if (string.IsNullOrWhiteSpace(rvtPath) || !File.Exists(rvtPath))
                    return;

                string targetFull = Path.GetFullPath(rvtPath);
                string targetFileNoExt = Path.GetFileNameWithoutExtension(targetFull);
                string targetFileName = Path.GetFileName(targetFull);
                string targetCentral = null;
                try
                {
                    var bi = BasicFileInfo.Extract(targetFull);
                    if (bi != null && bi.IsWorkshared) targetCentral = bi.CentralPath;
                }
                catch { }

                foreach (Document d in app.Documents)
                {
                    try
                    {
                        if (!string.IsNullOrEmpty(d.PathName))
                        {
                            string dFull = Path.GetFullPath(d.PathName);
                            if (string.Equals(dFull, targetFull, StringComparison.InvariantCultureIgnoreCase)) { doc = d; break; }
                            string dName = Path.GetFileName(d.PathName);
                            if (!string.IsNullOrEmpty(dName) &&
                                string.Equals(dName, targetFileName, StringComparison.InvariantCultureIgnoreCase)) { doc = d; break; }
                        }
                        if (!string.IsNullOrEmpty(d.Title) &&
                            string.Equals(d.Title, targetFileNoExt, StringComparison.InvariantCultureIgnoreCase)) { doc = d; break; }

                        if (!string.IsNullOrEmpty(targetCentral) && d.IsWorkshared)
                        {
                            string dCentral = null;
                            try
                            {
                                var mp = d.GetWorksharingCentralModelPath();
                                if (mp != null) dCentral = ModelPathUtils.ConvertModelPathToUserVisiblePath(mp);
                            }
                            catch { }
                            if (!string.IsNullOrEmpty(dCentral) &&
                                string.Equals(dCentral, targetCentral, StringComparison.InvariantCultureIgnoreCase)) { doc = d; break; }
                        }
                    }
                    catch { }
                }

                if (doc == null)
                {
                    var mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(rvtPath);
                    var openOpts = new OpenOptions();

                    if (!string.IsNullOrEmpty(targetCentral))
                    {
                        openOpts.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;
                    }

                    doc = app.OpenDocumentFile(mp, openOpts);
                    docOpenedHere = true;
                }

                if (doc == null)
                    throw new InvalidOperationException("Не удалось получить документ Revit.");

                var views = new FilteredElementCollector(doc).OfClass(typeof(ViewDrafting)).ToElements();
                var draftingViews = new List<ViewDrafting>();
                foreach (var e in views)
                {
                    var vd = e as ViewDrafting;
                    if (vd != null && !vd.IsTemplate) draftingViews.Add(vd);
                }

                using (var conn = new SQLiteConnection("Data Source=" + DbPath + ";Version=3;FailIfMissing=False;"))
                {
                    conn.Open();
                    EnsureSchema(conn);

                    // загружаем уже существующие пары "вид + файл"
                    var existing = LoadExistingByViewAndPath(conn);

                    using (var tx = conn.BeginTransaction())
                    {
                        foreach (var v in draftingViews)
                        {
                            string viewName = v.Name ?? string.Empty;
                            string scaleStr = "1:" + Math.Max(1, v.Scale);

                            var jo = new JObject
                            {
                                ["ИМЯ ВИДА"] = viewName,
                                ["МАСШТАБ"] = scaleStr
                            };
                            string json = JsonConvert.SerializeObject(jo, Formatting.None);

                            string now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                            string uName = userName ?? string.Empty;

                            // ключ по имени вида + пути к модели
                            string key = MakeViewKey(viewName, targetFull);

                            if (existing.TryGetValue(key, out long id))
                            {
                                using (var cmd = conn.CreateCommand())
                                {
                                    cmd.CommandText =
                                        "UPDATE nodeManager " +
                                        "SET PROP = @prop, " +
                                        "    LAST_SEEN = @run, " +
                                        "    RVT_PATH = @rvt " +
                                        "WHERE ID = @id;";

                                    cmd.Parameters.AddWithValue("@prop", json);
                                    cmd.Parameters.AddWithValue("@run", runToken);
                                    cmd.Parameters.AddWithValue("@rvt", targetFull);
                                    cmd.Parameters.AddWithValue("@id", id);

                                    cmd.ExecuteNonQuery();
                                }
                            }
                            else
                            {
                                using (var cmd = conn.CreateCommand())
                                {
                                    cmd.CommandText =
                                        "INSERT INTO nodeManager " +
                                        "(NAME, PROP, LAST_SEEN, RVT_PATH, TIME_CREATE, TIME_UPDATE, USER_CREATE, USER_UPDATE) " +
                                        "VALUES " +
                                        "(@name, @prop, @run, @rvt, @timeCr, @timeUpd, @userCr, @userUpd);";

                                    cmd.Parameters.AddWithValue("@name", viewName);
                                    cmd.Parameters.AddWithValue("@prop", json);
                                    cmd.Parameters.AddWithValue("@run", runToken);
                                    cmd.Parameters.AddWithValue("@rvt", targetFull);
                                    cmd.Parameters.AddWithValue("@timeCr", now);
                                    cmd.Parameters.AddWithValue("@timeUpd", now);
                                    cmd.Parameters.AddWithValue("@userCr", uName);
                                    cmd.Parameters.AddWithValue("@userUpd", uName);

                                    cmd.ExecuteNonQuery();
                                }
                            }

                        }

                        tx.Commit();
                    }
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("UpdateDb", "Ошибка при обновлении базы:\n" + ex.Message);
            }
            finally
            {
                if (docOpenedHere && doc != null)
                {
                    try { doc.Close(false); } catch { }
                }
            }
        }

        public static void FinalizeUpdateDb(string runToken, List<string> rvtPaths, bool isBimRun)
        {
            try
            {
                using (var conn = new SQLiteConnection("Data Source=" + DbPath + ";Version=3;FailIfMissing=False;"))
                {
                    conn.Open();
                    EnsureSchema(conn);

                    using (var tx = conn.BeginTransaction())
                    using (var cmd = conn.CreateCommand())
                    {
                        if (isBimRun)
                        {
                            cmd.CommandText = @"
                                DELETE FROM nodeManager
                                WHERE LAST_SEEN IS NULL OR LAST_SEEN <> @run;";
                            cmd.Parameters.AddWithValue("@run", runToken);
                            cmd.ExecuteNonQuery();
                        }
                        else
                        {
                            if (rvtPaths != null && rvtPaths.Count > 0)
                            {
                                var normalized = rvtPaths
                                    .Where(p => !string.IsNullOrWhiteSpace(p))
                                    .Select(p =>
                                    {
                                        var s = p.Trim();
                                        try { s = System.IO.Path.GetFullPath(s); }
                                        catch { }
                                        return s;
                                    })
                                    .Distinct(StringComparer.InvariantCultureIgnoreCase)
                                    .ToList();

                                if (normalized.Count > 0)
                                {
                                    var paramNames = new List<string>();
                                    for (int i = 0; i < normalized.Count; i++)
                                    {
                                        paramNames.Add("@p" + i);
                                    }

                                    cmd.CommandText = $@"
                                        DELETE FROM nodeManager
                                        WHERE (LAST_SEEN IS NULL OR LAST_SEEN <> @run)
                                          AND RVT_PATH IN ({string.Join(", ", paramNames)});";

                                    cmd.Parameters.AddWithValue("@run", runToken);

                                    for (int i = 0; i < normalized.Count; i++)
                                    {
                                        cmd.Parameters.AddWithValue("@p" + i, normalized[i]);
                                    }

                                    cmd.ExecuteNonQuery();
                                }
                            }
                        }

                        tx.Commit();
                    }
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("UpdateDb", "Ошибка при финализации базы:\n" + ex.Message);
            }
        }
    
        private static void EnsureSchema(SQLiteConnection conn)
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText =
                    "CREATE TABLE IF NOT EXISTS nodeManager (" +
                    "ID INTEGER PRIMARY KEY AUTOINCREMENT," +
                    "NAME TEXT NOT NULL," +
                    "PROP TEXT NOT NULL" +
                    ");";
                cmd.ExecuteNonQuery();
            }

            bool hasLastSeen = false;
            bool hasRvtPath = false;
            bool hasComment = false;
            bool hasTimeCreate = false;
            bool hasTimeUpdate = false;
            bool hasUserCreate = false;
            bool hasUserUpdate = false;

            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "PRAGMA table_info(nodeManager);";
                using (var r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        string col = r.GetString(1);
                        if (string.Equals(col, "LAST_SEEN", StringComparison.InvariantCultureIgnoreCase))
                            hasLastSeen = true;
                        else if (string.Equals(col, "RVT_PATH", StringComparison.InvariantCultureIgnoreCase))
                            hasRvtPath = true;
                        else if (string.Equals(col, "COMMENT", StringComparison.InvariantCultureIgnoreCase))
                            hasComment = true;
                        else if (string.Equals(col, "TIME_CREATE", StringComparison.InvariantCultureIgnoreCase))
                            hasTimeCreate = true;
                        else if (string.Equals(col, "TIME_UPDATE", StringComparison.InvariantCultureIgnoreCase))
                            hasTimeUpdate = true;
                        else if (string.Equals(col, "USER_CREATE", StringComparison.InvariantCultureIgnoreCase))
                            hasUserCreate = true;
                        else if (string.Equals(col, "USER_UPDATE", StringComparison.InvariantCultureIgnoreCase))
                            hasUserUpdate = true;
                    }
                }
            }

            using (var cmd = conn.CreateCommand())
            {
                if (!hasLastSeen)
                {
                    cmd.CommandText = "ALTER TABLE nodeManager ADD COLUMN LAST_SEEN TEXT;";
                    cmd.ExecuteNonQuery();
                }
                if (!hasRvtPath)
                {
                    cmd.CommandText = "ALTER TABLE nodeManager ADD COLUMN RVT_PATH TEXT;";
                    cmd.ExecuteNonQuery();
                }
                if (!hasComment)
                {
                    cmd.CommandText = "ALTER TABLE nodeManager ADD COLUMN COMMENT TEXT;";
                    cmd.ExecuteNonQuery();
                }
                if (!hasTimeCreate)
                {
                    cmd.CommandText = "ALTER TABLE nodeManager ADD COLUMN TIME_CREATE TEXT;";
                    cmd.ExecuteNonQuery();
                }
                if (!hasTimeUpdate)
                {
                    cmd.CommandText = "ALTER TABLE nodeManager ADD COLUMN TIME_UPDATE TEXT;";
                    cmd.ExecuteNonQuery();
                }
                if (!hasUserCreate)
                {
                    cmd.CommandText = "ALTER TABLE nodeManager ADD COLUMN USER_CREATE TEXT;";
                    cmd.ExecuteNonQuery();
                }
                if (!hasUserUpdate)
                {
                    cmd.CommandText = "ALTER TABLE nodeManager ADD COLUMN USER_UPDATE TEXT;";
                    cmd.ExecuteNonQuery();
                }
            }
        }

        private static Dictionary<string, long> LoadExistingByViewAndPath(SQLiteConnection conn)
        {
            var map = new Dictionary<string, long>(StringComparer.InvariantCultureIgnoreCase);

            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT ID, PROP, RVT_PATH FROM nodeManager;";
                using (var r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        try
                        {
                            long id = r.GetInt64(0);
                            string propJson = r.IsDBNull(1) ? null : r.GetString(1);
                            string rvtPath = r.IsDBNull(2) ? null : r.GetString(2);

                            if (string.IsNullOrWhiteSpace(propJson))
                                continue;

                            var jo = JObject.Parse(propJson);
                            string viewName = jo["ИМЯ ВИДА"]?.ToString();

                            if (string.IsNullOrWhiteSpace(viewName))
                                continue;

                            string key = MakeViewKey(viewName, rvtPath);
                            if (!map.ContainsKey(key))
                                map[key] = id;
                        }
                        catch
                        {
                        }
                    }
                }
            }

            return map;
        }

        private static string MakeViewKey(string viewName, string rvtPath)
        {
            string vn = (viewName ?? string.Empty).Trim();
            string path;

            try
            {
                path = string.IsNullOrWhiteSpace(rvtPath)
                    ? string.Empty
                    : System.IO.Path.GetFullPath(rvtPath);
            }
            catch
            {
                path = rvtPath ?? string.Empty;
            }

            return vn + "||" + path;
        }

        private void BtnCopyInView_Click(object sender, RoutedEventArgs e)
        {
            if (_currentElement == null)
                return;

            var uidoc = _uiapp?.ActiveUIDocument;

            _copyHandler.Init(_uiapp, DbPath, _currentElement.Id, this);
            _copyEvent.Raise();
        }

        private void BtnCopyElements_Click(object sender, RoutedEventArgs e)
        {
            if (_currentElement == null)
                return;

            var uidoc = _uiapp?.ActiveUIDocument;
            if (uidoc == null)
                return;

            var doc = uidoc.Document;
            if (doc == null)
                return;

            string propJson = null;
            string rvtPath = null;

            try
            {
                using (var conn = new System.Data.SQLite.SQLiteConnection("Data Source=" + DbPath + ";Version=3;FailIfMissing=False;"))
                {
                    conn.Open();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "SELECT PROP, RVT_PATH FROM nodeManager WHERE ID = @id;";
                        cmd.Parameters.AddWithValue("@id", _currentElement.Id);

                        using (var r = cmd.ExecuteReader())
                        {
                            if (r.Read())
                            {
                                if (!r.IsDBNull(0))
                                    propJson = r.GetString(0);

                                if (!r.IsDBNull(1))
                                    rvtPath = r.GetString(1);
                            }
                        }
                    }
                }
            }


            catch (Exception ex)
            {
                TaskDialog.Show("KPLN. Менеджер узлов", "Ошибка при чтении базы данных:\n" + ex.Message);
                return;
            }

            if (string.IsNullOrWhiteSpace(propJson))
            {
                TaskDialog.Show("KPLN. Менеджер узлов", "Для выбранного узла отсутствует PROP.");
                return;
            }

            string sourceViewName = null;
            try
            {
                var jo = Newtonsoft.Json.Linq.JObject.Parse(propJson);
                sourceViewName = jo["ИМЯ ВИДА"]?.ToString();
            }
            catch { }

            if (string.IsNullOrWhiteSpace(sourceViewName))
            {
                TaskDialog.Show("KPLN. Менеджер узлов", "В PROP нет ключа \"ИМЯ ВИДА\".");
                return;
            }

            if (string.IsNullOrWhiteSpace(rvtPath))
            {
                TaskDialog.Show("KPLN. Менеджер узлов", "Для выбранного узла отсутствует RVT_PATH.");
                return;
            }

            var activeView = uidoc.ActiveView;

            bool isLegend = activeView.ViewType == ViewType.Legend;

            if (!(activeView is ViewSheet) && !(activeView is ViewDrafting) && !isLegend)
            {
                TaskDialog.Show("KPLN. Менеджер узлов", "Копирование узла поддерживается только, если открыт лист, чертёжный вид или легенда.");
                return;
            }

            _placeViewHandler.Init(_uiapp, rvtPath, sourceViewName, this);
            _placeViewEvent.Raise();
        }
    }































    /////////////////////////////// КОПИРОВАНИЕ ВИДОВ
    internal sealed class UseDestinationTypesHandler : IDuplicateTypeNamesHandler
    {
        public DuplicateTypeAction OnDuplicateTypeNamesFound(DuplicateTypeNamesHandlerArgs args)
            => DuplicateTypeAction.UseDestinationTypes;
    }

    internal sealed class CopyNodeToViewHandler : IExternalEventHandler
    {
        private UIApplication _uiapp;
        private string _dbPath;
        private long _nodeId;

        private Window _ownerWindow;

        public void Init(UIApplication uiapp, string dbPath, long nodeId, Window ownerWindow)
        {
            _uiapp = uiapp;
            _dbPath = dbPath;
            _nodeId = nodeId;

            _ownerWindow = ownerWindow;
        }

        public string GetName() => "KPLN. Копирование узла";

        private static string TryGetDwgPathFromInstance(Document doc, ImportInstance inst)
        {
            var linkType = doc.GetElement(inst.GetTypeId()) as CADLinkType;
            if (linkType == null)
                return null;

            var extRef = ExternalFileUtils.GetExternalFileReference(doc, linkType.Id);
            if (extRef == null)
                return null;

            var modelPath = extRef.GetPath();
            if (modelPath == null)
                return null;

            string userPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath);
            return userPath;
        }

        private static string MakeSafeFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "DWG";

            var invalid = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder(name.Length);
            foreach (var ch in name)
            {
                sb.Append(invalid.Contains(ch) ? '_' : ch);
            }
            return sb.ToString();
        }

        private static string GetTargetBaseModelPath(Document targetDoc)
        {
            try
            {
                if (targetDoc.IsWorkshared)
                {
                    var centralMp = targetDoc.GetWorksharingCentralModelPath();
                    if (centralMp != null)
                        return ModelPathUtils.ConvertModelPathToUserVisiblePath(centralMp);
                }
            }
            catch { }
            return targetDoc.PathName;
        }

        private static void Apply2DTransformLikeSource(Document doc, ElementId importInstanceId, Transform srcTr)
        {
            if (srcTr == null)
                return;

            XYZ move = srcTr.Origin ?? XYZ.Zero;
            if (move != null && !move.IsZeroLength())
                ElementTransformUtils.MoveElement(doc, importInstanceId, move);

            XYZ bx = srcTr.BasisX ?? XYZ.BasisX;
            double angle = Math.Atan2(bx.Y, bx.X);

            if (Math.Abs(angle) > 1e-9)
            {
                XYZ p = srcTr.Origin ?? XYZ.Zero;
                var axis = Line.CreateBound(p, p + XYZ.BasisZ);
                ElementTransformUtils.RotateElement(doc, importInstanceId, axis, angle);
            }
        }

        public void Execute(UIApplication app)
        {
            try
            {
                var uidoc = _uiapp?.ActiveUIDocument ?? throw new InvalidOperationException("Нет активного документа.");
                var targetDoc = uidoc.Document;

                string rvtPath = null, propJson = null;
                using (var conn = new System.Data.SQLite.SQLiteConnection("Data Source=" + _dbPath + ";Version=3;FailIfMissing=False;"))
                {
                    conn.Open();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "SELECT RVT_PATH, PROP FROM nodeManager WHERE ID = @id;";
                        cmd.Parameters.AddWithValue("@id", _nodeId);
                        using (var r = cmd.ExecuteReader())
                        {
                            if (r.Read())
                            {
                                if (!r.IsDBNull(0)) rvtPath = r.GetString(0);
                                if (!r.IsDBNull(1)) propJson = r.GetString(1);
                            }
                        }
                    }
                }

                if (string.IsNullOrWhiteSpace(rvtPath) || !File.Exists(rvtPath))
                    throw new FileNotFoundException("RVT_PATH не найден.", rvtPath ?? "");
                if (string.IsNullOrWhiteSpace(propJson))
                    throw new InvalidOperationException("PROP отсутствует.");

                string sourceViewName = null;
                try
                {
                    var jo = Newtonsoft.Json.Linq.JObject.Parse(propJson);
                    sourceViewName = jo["ИМЯ ВИДА"]?.ToString();
                }
                catch { }
                if (string.IsNullOrWhiteSpace(sourceViewName))
                    throw new InvalidOperationException("В PROP нет ключа \"ИМЯ ВИДА\".");

                Document sourceDoc = _uiapp.Application.Documents
                    .Cast<Document>()
                    .FirstOrDefault(d =>
                    {
                        try
                        {
                            if (string.IsNullOrEmpty(d.PathName)) return false;
                            return string.Equals(Path.GetFullPath(d.PathName), Path.GetFullPath(rvtPath), StringComparison.InvariantCultureIgnoreCase);
                        }
                        catch { return false; }
                    });

                bool openedSourceHere = false;
                if (sourceDoc == null)
                {
                    sourceDoc = _uiapp.Application.OpenDocumentFile(rvtPath);
                    openedSourceHere = true;
                }

                try
                {
                    if (string.Equals(Path.GetFullPath(sourceDoc.PathName), Path.GetFullPath(targetDoc.PathName), StringComparison.InvariantCultureIgnoreCase))
                        throw new InvalidOperationException("Нельзя копировать в тот же документ.");
                }
                catch { }


                var sourceView = new FilteredElementCollector(sourceDoc)
                    .OfClass(typeof(View)).Cast<View>().FirstOrDefault(v => !v.IsTemplate && string.Equals(v.Name, sourceViewName, StringComparison.InvariantCultureIgnoreCase))
                    ?? throw new InvalidOperationException($"В документе-источнике не найден вид \"{sourceViewName}\".");

                var allInView = new FilteredElementCollector(sourceDoc, sourceView.Id).WhereElementIsNotElementType().ToElements();

                // DWG
                bool hasAnyDwg = allInView.OfType<ImportInstance>().Any();
                if (hasAnyDwg)
                {
                    // Проверяем: есть ли на исходном виде что-то кроме DWG 
                    bool hasNonDwgStuff = new FilteredElementCollector(sourceDoc, sourceView.Id)
                        .WhereElementIsNotElementType().Any(e =>
                            e.ViewSpecific && !(e is ImportInstance) && !(e is View) && !(e is Group) && e.Category != null);

                    // Есть DWG + есть аннотации/прочее
                    // ============================================================
                    if (hasNonDwgStuff)
                    {


                        TaskDialog.Show("Менеджер узлов","Вы пытаетесь добавить узел, который состоит одновремено из DWG и встроенной графики. " +
                            "Добавление таких узлов запрещено: обратитесь к разработчику узла для того, чтобы он поправил данный узел.");
                        return;


                        string baseModelPath = GetTargetBaseModelPath(targetDoc);
                        string baseSafeName = MakeSafeFileName(sourceViewName);

                        string dwgFolder = null;
                        bool canUseProjectFolder = !string.IsNullOrWhiteSpace(baseModelPath) && File.Exists(baseModelPath);

                        if (!canUseProjectFolder)
                        {
                            string mainDb = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_Loader_MainDB.db";

                            string foundMainPath = null;
                            int? foundSubDepId = null;
                            string foundSubDepCode = null;

                            try
                            {
                                using (var conn = new System.Data.SQLite.SQLiteConnection(
                                    "Data Source=" + mainDb + ";Version=3;FailIfMissing=True;"))
                                {
                                    conn.Open();

                                    using (var cmd = conn.CreateCommand())
                                    {
                                        cmd.CommandText = @"
                                            SELECT MainPath, RevitServerPath, RevitServerPath2, RevitServerPath3, RevitServerPath4
                                            FROM Projects
                                            WHERE (RevitServerPath  IS NOT NULL AND TRIM(RevitServerPath)  <> '')
                                               OR (RevitServerPath2 IS NOT NULL AND TRIM(RevitServerPath2) <> '')
                                               OR (RevitServerPath3 IS NOT NULL AND TRIM(RevitServerPath3) <> '')
                                               OR (RevitServerPath4 IS NOT NULL AND TRIM(RevitServerPath4) <> '');";

                                        using (var r = cmd.ExecuteReader())
                                        {
                                            while (r.Read())
                                            {
                                                string mp = r.IsDBNull(0) ? null : r.GetString(0);

                                                string p1 = r.IsDBNull(1) ? null : r.GetString(1);
                                                string p2 = r.IsDBNull(2) ? null : r.GetString(2);
                                                string p3 = r.IsDBNull(3) ? null : r.GetString(3);
                                                string p4 = r.IsDBNull(4) ? null : r.GetString(4);

                                                bool match =
                                                    (!string.IsNullOrWhiteSpace(p1) && baseModelPath.StartsWith(p1.Trim(), StringComparison.InvariantCultureIgnoreCase)) ||
                                                    (!string.IsNullOrWhiteSpace(p2) && baseModelPath.StartsWith(p2.Trim(), StringComparison.InvariantCultureIgnoreCase)) ||
                                                    (!string.IsNullOrWhiteSpace(p3) && baseModelPath.StartsWith(p3.Trim(), StringComparison.InvariantCultureIgnoreCase)) ||
                                                    (!string.IsNullOrWhiteSpace(p4) && baseModelPath.StartsWith(p4.Trim(), StringComparison.InvariantCultureIgnoreCase));

                                                if (match)
                                                {
                                                    foundMainPath = mp;
                                                    break;
                                                }
                                            }
                                        }
                                    }

                                    if (string.IsNullOrWhiteSpace(foundMainPath))
                                    {
                                        TaskDialog.Show("KPLN. Менеджер узлов",
                                            "В MainDB в таблице Projects не найдено соответствий RevitServerPath* и MainPath.\n" +
                                            "Путь к файлу на Revit-сервере:\n" + (baseModelPath ?? "<null>") + "\n" +
                                            "Для решения проблемы - обратитесь в BIM-отдел");
                                        return;
                                    }

                                    using (var cmd = conn.CreateCommand())
                                    {
                                        cmd.CommandText = @"
                                            SELECT SubDepartmentId
                                            FROM Documents
                                            WHERE CentralPath = @p
                                            LIMIT 1;";
                                        cmd.Parameters.AddWithValue("@p", baseModelPath);

                                        var obj = cmd.ExecuteScalar();
                                        if (obj != null && obj != DBNull.Value && int.TryParse(obj.ToString(), out int dep))
                                            foundSubDepId = dep;
                                    }

                                    if (!foundSubDepId.HasValue)
                                    {
                                        TaskDialog.Show("KPLN. Менеджер узлов",
                                            "Файл с данным CentralPath не зарегистрирован в БД (таблица Documents).\n" +
                                            "Текущий CentralPath:\n" + (baseModelPath ?? "<null>") + "\n" +
                                            "Для решения проблемы - обратитесь в BIM-отдел");
                                        return;
                                    }

                                    using (var cmd = conn.CreateCommand())
                                    {
                                        cmd.CommandText = @"
                                            SELECT Code
                                            FROM SubDepartments
                                            WHERE Id = @id
                                            LIMIT 1;";
                                        cmd.Parameters.AddWithValue("@id", foundSubDepId.Value);

                                        var obj = cmd.ExecuteScalar();
                                        foundSubDepCode = (obj == null || obj == DBNull.Value) ? null : obj.ToString();
                                        if (!string.IsNullOrWhiteSpace(foundSubDepCode))
                                            foundSubDepCode = foundSubDepCode.Trim();
                                    }

                                    conn.Close();
                                }
                            }
                            catch (Exception ex)
                            {
                                TaskDialog.Show("KPLN. Менеджер узлов", "Ошибка чтения MainDB:\n" + ex.Message);
                                return;
                            }

                            string mainPathDisplay = foundMainPath.Replace(@"\\stinproject.local\project\", @"Y:\");

                            string depDisplay = !string.IsNullOrWhiteSpace(foundSubDepCode)
                                ? foundSubDepCode
                                : foundSubDepId.Value.ToString(CultureInfo.InvariantCulture);

                            // Формируем dwgFolder = <mainPathDisplay>\BIM\8.Менеджер узлов\<depDisplay> ---
                            try
                            {
                                foreach (char ch in Path.GetInvalidFileNameChars())
                                    depDisplay = depDisplay.Replace(ch.ToString(), "_");

                                string baseBim = Path.Combine(mainPathDisplay, "BIM");
                                string managerDir = Path.Combine(baseBim, "8.Менеджер узлов");
                                dwgFolder = Path.Combine(managerDir, depDisplay);

                                Directory.CreateDirectory(dwgFolder);
                            }
                            catch (Exception ex)
                            {
                                TaskDialog.Show("KPLN. Менеджер узлов", "Не удалось создать директорию для DWG:\n" +
                                    (dwgFolder ?? "<null>") + "\n\n" + ex.Message);
                                return;
                            }

                            // Если хочешь показать пользователю — оставь, иначе можно убрать
                            TaskDialog.Show("KPLN. Менеджер узлов", "Вы работаете с файлом на Revit-сервере. DWG будет сохранён в:\n" + dwgFolder);
                        }

                        else
                        {
                            string modelDir = Path.GetDirectoryName(baseModelPath);
                            dwgFolder = Path.Combine(modelDir, "DWG_Менеджер узлов");
                            Directory.CreateDirectory(dwgFolder);
                        }

                        ViewDrafting targetView;
                        using (var t = new Transaction(targetDoc, "KPLN. Поиск/создание и настройка вида узла"))
                        {
                            t.Start();

                            targetView = new FilteredElementCollector(targetDoc)
                                .OfClass(typeof(ViewDrafting))
                                .Cast<ViewDrafting>()
                                .FirstOrDefault(v => !v.IsTemplate &&
                                    string.Equals(v.Name, sourceViewName, StringComparison.InvariantCultureIgnoreCase));

                            if (targetView == null)
                            {
                                var draftingTypeId = new FilteredElementCollector(targetDoc)
                                    .OfClass(typeof(ViewFamilyType))
                                    .Cast<ViewFamilyType>()
                                    .First(vft => vft.ViewFamily == ViewFamily.Drafting)
                                    .Id;

                                targetView = ViewDrafting.Create(targetDoc, draftingTypeId);
                                targetView.Name = sourceViewName;
                            }

                            try { targetView.Scale = sourceView.Scale; } catch { }
                            try { targetView.Discipline = sourceView.Discipline; } catch { }
                            try { targetView.DetailLevel = sourceView.DetailLevel; } catch { }
                            try { targetView.DisplayStyle = sourceView.DisplayStyle; } catch { }
                            try { targetView.PartsVisibility = sourceView.PartsVisibility; } catch { }

                            try
                            {
                                if (sourceView.ViewTemplateId != ElementId.InvalidElementId)
                                {
                                    var srcTemplate = sourceDoc.GetElement(sourceView.ViewTemplateId) as View;
                                    if (srcTemplate != null)
                                    {
                                        string templateName = srcTemplate.Name;

                                        var dstTemplate = new FilteredElementCollector(targetDoc)
                                            .OfClass(typeof(View))
                                            .Cast<View>()
                                            .FirstOrDefault(v => v.IsTemplate &&
                                                string.Equals(v.Name, templateName, StringComparison.InvariantCultureIgnoreCase));

                                        if (dstTemplate != null)
                                        {
                                            targetView.ViewTemplateId = dstTemplate.Id;
                                        }
                                        else
                                        {
                                            TaskDialog.Show("KPLN. Менеджер узлов",
                                                $"В исходном документе у вида установлен шаблон: \"{templateName}\"\n" +
                                                "Но в текущем проекте шаблон вида с таким именем не найден.\n" +
                                                "Вид будет создан/обновлён без назначения шаблона. Возможно некорректное отображение DWG/аннотаций.");
                                        }
                                    }
                                }
                            }
                            catch { }

                            t.Commit();
                        }

                        using (var t = new Transaction(targetDoc, "KPLN. Очистка вида узла перед вставкой слепка DWG"))
                        {
                            t.Start();

                            var toDelete = new FilteredElementCollector(targetDoc, targetView.Id)
                                .WhereElementIsNotElementType()
                                .Where(e =>
                                    e.ViewSpecific &&
                                    !(e is View) &&
                                    !(e is Group) &&
                                    e.Category != null)
                                .Select(e => e.Id)
                                .ToList();

                            if (toDelete.Count > 0)
                                targetDoc.Delete(toDelete);

                            t.Commit();
                        }

                        string snapshotNameNoExt = baseSafeName;
                        string snapshotPath = Path.Combine(dwgFolder, snapshotNameNoExt + ".dwg");

                        if (File.Exists(snapshotPath))
                        {
                            var msg =
                                $"Файл DWG уже существует:\n{snapshotPath}\n\n" +
                                "Да — перезаписать DWG и перезаписать содержимое вида.\n" +
                                "Нет — открыть существующий вид.\n" +
                                "Отмена — ничего не менять.";

                            var result = System.Windows.MessageBox.Show(
                                msg, "KPLN. Менеджер узлов",
                                MessageBoxButton.YesNoCancel, MessageBoxImage.Question);

                            if (result == MessageBoxResult.Cancel)
                            {
                                if (openedSourceHere) { try { sourceDoc.Close(false); } catch { } }
                                return;
                            }
                            else if (result == MessageBoxResult.No)
                            {
                                uidoc.ActiveView = targetView;
                                if (openedSourceHere) { try { sourceDoc.Close(false); } catch { } }
                                return;
                            }
                            else
                            {
                                try { File.Delete(snapshotPath); } catch { }
                            }
                        }

                        bool exportOk = false;
                        string exportErr = null;
                        try
                        {
                            var viewIds = new List<ElementId> { sourceView.Id };
                            var opt = new DWGExportOptions();

                            exportOk = sourceDoc.Export(dwgFolder, snapshotNameNoExt, viewIds, opt);
                        }
                        catch (Exception ex)
                        {
                            exportErr = ex.Message;
                        }

                        if (!exportOk || !File.Exists(snapshotPath))
                        {
                            TaskDialog.Show("KPLN. Менеджер узлов", "Не удалось экспортировать вид в DWG.\n" +
                                (exportErr != null ? ("\n" + exportErr) : ""));

                            if (openedSourceHere) { try { sourceDoc.Close(false); } catch { } }
                            return;
                        }

                        using (var t = new Transaction(targetDoc, "KPLN. Линковка слепка DWG"))
                        {
                            t.Start();

                            var opt = new DWGImportOptions
                            {
                                ThisViewOnly = true,
                                Placement = ImportPlacement.Origin,
                                OrientToView = true
                            };

                            ElementId linkedId;
                            bool ok = targetDoc.Link(snapshotPath, opt, targetView, out linkedId);

                            t.Commit();

                            if (!ok || linkedId == ElementId.InvalidElementId)
                            {
                                TaskDialog.Show("KPLN. Менеджер узлов",
                                    "DWG экспортирован, но не удалось залинковать его в целевой вид.");

                                if (openedSourceHere) { try { sourceDoc.Close(false); } catch { } }
                                return;
                            }
                        }

                        if (openedSourceHere) { try { sourceDoc.Close(false); } catch { } }

                        uidoc.ActiveView = targetView;

                        TaskDialog.Show("KPLN. Менеджер узлов", $"Готово");
                        return;
                    }

                    // Чисто DWG 
                    // ============================================================

                    var dwgInstances = allInView.OfType<ImportInstance>().ToList();
                    var dwgPlacements = new List<(ImportInstance inst, string fullPath, Autodesk.Revit.DB.Transform tr)>();

                    foreach (var inst in dwgInstances)
                    {
                        string path = TryGetDwgPathFromInstance(sourceDoc, inst);
                        if (string.IsNullOrWhiteSpace(path))
                            continue;

                        try { path = Path.GetFullPath(path); } catch { }

                        if (!File.Exists(path))
                            continue;

                        Autodesk.Revit.DB.Transform tr = Transform.Identity;
                        try { tr = inst.GetTransform(); } catch { }

                        dwgPlacements.Add((inst, path, tr));
                    }

                    if (dwgPlacements.Count == 0)
                    {
                        TaskDialog.Show("KPLN. Менеджер узлов",
                            "На исходном виде обнаружен DWG, но это, вероятно, импорт без живой ссылки. " +
                            "Перенос такого DWG в другой документ не поддерживается Revit API.");

                        if (openedSourceHere) { try { sourceDoc.Close(false); } catch { } }
                        return;
                    }

                    string baseModelPath2 = GetTargetBaseModelPath(targetDoc);
                    string baseSafeName2 = MakeSafeFileName(sourceViewName);

                    string dwgFolder2 = null;
                    bool canUseProjectFolder2 = !string.IsNullOrWhiteSpace(baseModelPath2) && File.Exists(baseModelPath2);

                    if (!canUseProjectFolder2)
                    {
                        string mainDb = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_Loader_MainDB.db";

                        string foundMainPath = null;
                        int? foundSubDepId = null;
                        string foundSubDepCode = null;

                        try
                        {
                            using (var conn = new System.Data.SQLite.SQLiteConnection(
                                "Data Source=" + mainDb + ";Version=3;FailIfMissing=True;"))
                            {
                                conn.Open();

                                using (var cmd = conn.CreateCommand())
                                {
                                    cmd.CommandText = @"
                                            SELECT MainPath, RevitServerPath, RevitServerPath2, RevitServerPath3, RevitServerPath4
                                            FROM Projects
                                            WHERE (RevitServerPath  IS NOT NULL AND TRIM(RevitServerPath)  <> '')
                                               OR (RevitServerPath2 IS NOT NULL AND TRIM(RevitServerPath2) <> '')
                                               OR (RevitServerPath3 IS NOT NULL AND TRIM(RevitServerPath3) <> '')
                                               OR (RevitServerPath4 IS NOT NULL AND TRIM(RevitServerPath4) <> '');";

                                    using (var r = cmd.ExecuteReader())
                                    {
                                        while (r.Read())
                                        {
                                            string mp = r.IsDBNull(0) ? null : r.GetString(0);

                                            string p1 = r.IsDBNull(1) ? null : r.GetString(1);
                                            string p2 = r.IsDBNull(2) ? null : r.GetString(2);
                                            string p3 = r.IsDBNull(3) ? null : r.GetString(3);
                                            string p4 = r.IsDBNull(4) ? null : r.GetString(4);

                                            bool match =
                                                (!string.IsNullOrWhiteSpace(p1) && baseModelPath2.StartsWith(p1.Trim(), StringComparison.InvariantCultureIgnoreCase)) ||
                                                (!string.IsNullOrWhiteSpace(p2) && baseModelPath2.StartsWith(p2.Trim(), StringComparison.InvariantCultureIgnoreCase)) ||
                                                (!string.IsNullOrWhiteSpace(p3) && baseModelPath2.StartsWith(p3.Trim(), StringComparison.InvariantCultureIgnoreCase)) ||
                                                (!string.IsNullOrWhiteSpace(p4) && baseModelPath2.StartsWith(p4.Trim(), StringComparison.InvariantCultureIgnoreCase));

                                            if (match)
                                            {
                                                foundMainPath = mp;
                                                break;
                                            }
                                        }
                                    }
                                }

                                if (string.IsNullOrWhiteSpace(foundMainPath))
                                {
                                    TaskDialog.Show("KPLN. Менеджер узлов",
                                        "В MainDB в таблице Projects не найдено соответствий RevitServerPath* и MainPath.\n" +
                                        "Путь к файлу на Revit-сервере:\n" + (baseModelPath2 ?? "<null>") + "\n" +
                                        "Для решения проблемы - обратитесь в BIM-отдел");
                                    return;
                                }

                                using (var cmd = conn.CreateCommand())
                                {
                                    cmd.CommandText = @"
                                            SELECT SubDepartmentId
                                            FROM Documents
                                            WHERE CentralPath = @p
                                            LIMIT 1;";
                                    cmd.Parameters.AddWithValue("@p", baseModelPath2);

                                    var obj = cmd.ExecuteScalar();
                                    if (obj != null && obj != DBNull.Value && int.TryParse(obj.ToString(), out int dep))
                                        foundSubDepId = dep;
                                }

                                if (!foundSubDepId.HasValue)
                                {
                                    TaskDialog.Show("KPLN. Менеджер узлов",
                                        "Файл с данным CentralPath не зарегистрирован в БД (таблица Documents).\n" +
                                        "Текущий CentralPath:\n" + (baseModelPath2 ?? "<null>") + "\n" +
                                        "Для решения проблемы - обратитесь в BIM-отдел");
                                    return;
                                }

                                using (var cmd = conn.CreateCommand())
                                {
                                    cmd.CommandText = @"
                                            SELECT Code
                                            FROM SubDepartments
                                            WHERE Id = @id
                                            LIMIT 1;";
                                    cmd.Parameters.AddWithValue("@id", foundSubDepId.Value);

                                    var obj = cmd.ExecuteScalar();
                                    foundSubDepCode = (obj == null || obj == DBNull.Value) ? null : obj.ToString();
                                    if (!string.IsNullOrWhiteSpace(foundSubDepCode))
                                        foundSubDepCode = foundSubDepCode.Trim();
                                }

                                conn.Close();
                            }
                        }
                        catch (Exception ex)
                        {
                            TaskDialog.Show("KPLN. Менеджер узлов", "Ошибка чтения MainDB:\n" + ex.Message);
                            return;
                        }

                        string mainPathDisplay = foundMainPath.Replace(@"\\stinproject.local\project\", @"Y:\");

                        string depDisplay = !string.IsNullOrWhiteSpace(foundSubDepCode)
                            ? foundSubDepCode
                            : foundSubDepId.Value.ToString(CultureInfo.InvariantCulture);

                        try
                        {
                            foreach (char ch in Path.GetInvalidFileNameChars())
                                depDisplay = depDisplay.Replace(ch.ToString(), "_");

                            string baseBim = Path.Combine(mainPathDisplay, "BIM");
                            string managerDir = Path.Combine(baseBim, "8.Менеджер узлов");
                            dwgFolder2 = Path.Combine(managerDir, depDisplay);

                            Directory.CreateDirectory(dwgFolder2);
                        }
                        catch (Exception ex)
                        {
                            TaskDialog.Show("KPLN. Менеджер узлов", "Не удалось создать директорию для DWG:\n" +
                                (dwgFolder2 ?? "<null>") + "\n\n" + ex.Message);
                            return;
                        }

                        TaskDialog.Show("KPLN. Менеджер узлов", "Вы работаете с файлом на Revit-сервере. DWG будет сохранён в:\n" + dwgFolder2);                  
                }
                    else
                    {
                        string modelDir = Path.GetDirectoryName(baseModelPath2);
                        dwgFolder2 = Path.Combine(modelDir, "DWG_Менеджер узлов");
                        Directory.CreateDirectory(dwgFolder2);
                    }

                    var plannedLocalCopies = new List<(string srcFullPath, string localCopyPath)>();
                    for (int i = 0; i < dwgPlacements.Count; i++)
                    {
                        string fileName = (dwgPlacements.Count == 1)
                            ? (baseSafeName2 + ".dwg")
                            : (baseSafeName2 + "_" + (i + 1) + ".dwg");

                        string localCopy = Path.Combine(dwgFolder2, fileName);
                        plannedLocalCopies.Add((dwgPlacements[i].fullPath, localCopy));
                    }

                    bool anyExists = plannedLocalCopies.Any(x => File.Exists(x.localCopyPath));

                    if (anyExists)
                    {
                        string msg =
                            $"В папке уже существуют DWG-файлы для узла \"{sourceViewName}\":\n" +
                            $"{dwgFolder2}\n\n" +
                            "Да — перезаписать DWG (все) и перезаписать содержимое соответствующего вида узла.\n" +
                            "Нет — открыть существующий вид с этим узлом.\n" +
                            "Отмена — ничего не менять.";

                        var result = System.Windows.MessageBox.Show(
                            msg, "KPLN. Менеджер узлов",
                            MessageBoxButton.YesNoCancel, MessageBoxImage.Question);

                        if (result == MessageBoxResult.Cancel)
                        {
                            if (openedSourceHere) { try { sourceDoc.Close(false); } catch { } }
                            return;
                        }
                        else if (result == MessageBoxResult.No)
                        {
                            var existingView = new FilteredElementCollector(targetDoc)
                                .OfClass(typeof(ViewDrafting))
                                .Cast<ViewDrafting>()
                                .FirstOrDefault(v => !v.IsTemplate &&
                                    string.Equals(v.Name, sourceViewName, StringComparison.InvariantCultureIgnoreCase));

                            if (existingView != null)
                            {
                                uidoc.ActiveView = existingView;
                                if (openedSourceHere) { try { sourceDoc.Close(false); } catch { } }
                                return;
                            }
                            else
                            {
                                var ask = System.Windows.MessageBox.Show(
                                    $"DWG уже существует в папке,\nно вид \"{sourceViewName}\" в документе не найден.\n\n" +
                                    "Создать новый чертёжный вид с этим именем\nи залинковать на него существующие DWG?",
                                    "KPLN. Менеджер узлов",
                                    MessageBoxButton.YesNo, MessageBoxImage.Question);

                                if (ask == MessageBoxResult.No)
                                {
                                    if (openedSourceHere) { try { sourceDoc.Close(false); } catch { } }
                                    return;
                                }
                            }
                        }
                    }

                    for (int i = 0; i < plannedLocalCopies.Count; i++)
                    {
                        string src = plannedLocalCopies[i].srcFullPath;
                        string dst = plannedLocalCopies[i].localCopyPath;

                        try
                        {
                            File.Copy(src, dst, true);
                        }
                        catch (Exception ex)
                        {
                            TaskDialog.Show("KPLN. Менеджер узлов",
                                $"Не удалось скопировать DWG:\n{src}\n→ {dst}\n\n{ex.Message}");

                            if (openedSourceHere) { try { sourceDoc.Close(false); } catch { } }
                            return;
                        }
                    }

                    ViewDrafting targetView2;
                    using (var t = new Transaction(targetDoc, "KPLN. Поиск/создание и настройка вида узла"))
                    {
                        t.Start();

                        targetView2 = new FilteredElementCollector(targetDoc)
                            .OfClass(typeof(ViewDrafting))
                            .Cast<ViewDrafting>()
                            .FirstOrDefault(v => !v.IsTemplate &&
                                string.Equals(v.Name, sourceViewName, StringComparison.InvariantCultureIgnoreCase));

                        if (targetView2 == null)
                        {
                            var draftingTypeId = new FilteredElementCollector(targetDoc)
                                .OfClass(typeof(ViewFamilyType))
                                .Cast<ViewFamilyType>()
                                .First(vft => vft.ViewFamily == ViewFamily.Drafting)
                                .Id;

                            targetView2 = ViewDrafting.Create(targetDoc, draftingTypeId);
                            targetView2.Name = sourceViewName;
                        }

                        try { targetView2.Scale = sourceView.Scale; } catch { }
                        try { targetView2.Discipline = sourceView.Discipline; } catch { }
                        try { targetView2.DetailLevel = sourceView.DetailLevel; } catch { }
                        try { targetView2.DisplayStyle = sourceView.DisplayStyle; } catch { }
                        try { targetView2.PartsVisibility = sourceView.PartsVisibility; } catch { }

                        try
                        {
                            if (sourceView.ViewTemplateId != ElementId.InvalidElementId)
                            {
                                var srcTemplate = sourceDoc.GetElement(sourceView.ViewTemplateId) as View;
                                if (srcTemplate != null)
                                {
                                    string templateName = srcTemplate.Name;

                                    var dstTemplate = new FilteredElementCollector(targetDoc)
                                        .OfClass(typeof(View))
                                        .Cast<View>()
                                        .FirstOrDefault(v => v.IsTemplate &&
                                            string.Equals(v.Name, templateName, StringComparison.InvariantCultureIgnoreCase));

                                    if (dstTemplate != null)
                                    {
                                        targetView2.ViewTemplateId = dstTemplate.Id;
                                    }
                                    else
                                    {
                                        TaskDialog.Show("KPLN. Менеджер узлов",
                                            $"В исходном документе у вида установлен шаблон: \"{templateName}\"\n" +
                                            "Но в текущем проекте шаблон вида с таким именем не найден.\n" +
                                            "Вид будет создан/обновлён без назначения шаблона. Возможно некорректное отображение DWG.");
                                    }
                                }
                            }
                        }
                        catch { }

                        t.Commit();
                    }

                    using (var t = new Transaction(targetDoc, "KPLN. Очистка аннотации на виде узла"))
                    {
                        t.Start();

                        var toDelete = new FilteredElementCollector(targetDoc, targetView2.Id)
                            .WhereElementIsNotElementType()
                            .Where(e =>
                                e.ViewSpecific &&
                                !(e is ImportInstance) &&
                                !(e is View) &&
                                !(e is Group) &&
                                e.Category != null)
                            .Select(e => e.Id)
                            .ToList();

                        if (toDelete.Count > 0)
                            targetDoc.Delete(toDelete);

                        t.Commit();
                    }

                    using (var t = new Transaction(targetDoc, "KPLN. Очистка старых DWG на виде узла"))
                    {
                        t.Start();

                        var oldDwgs = new FilteredElementCollector(targetDoc, targetView2.Id)
                            .OfClass(typeof(ImportInstance))
                            .ToElementIds();

                        if (oldDwgs.Count > 0)
                            targetDoc.Delete(oldDwgs);

                        t.Commit();
                    }

                    using (var t = new Transaction(targetDoc, "KPLN. Линковка DWG с восстановлением позиции"))
                    {
                        t.Start();

                        for (int i = 0; i < dwgPlacements.Count; i++)
                        {
                            var srcPlacement = dwgPlacements[i];
                            string localCopyPath = plannedLocalCopies[i].localCopyPath;

                            var opt = new DWGImportOptions
                            {
                                ThisViewOnly = true,
                                Placement = ImportPlacement.Origin,
                                OrientToView = true
                            };

                            ElementId linkedId;
                            bool ok = targetDoc.Link(localCopyPath, opt, targetView2, out linkedId);
                            if (!ok || linkedId == ElementId.InvalidElementId)
                                continue;

                            Apply2DTransformLikeSource(targetDoc, linkedId, srcPlacement.tr);
                        }

                        t.Commit();
                    }

                    if (openedSourceHere)
                    {
                        try { sourceDoc.Close(false); } catch { }
                    }

                    uidoc.ActiveView = targetView2;

                    TaskDialog.Show("KPLN. Менеджер узлов",
                        $"DWG с вида \"{sourceViewName}\" скопирован(ы) в папку:\n\"{dwgFolder2}\"\nи залинкован(ы) на вид в активном документе.");

                    return;
                }
                // ВСЁ ОСТАЛЬНОЕ
                else
                {
                    ViewDrafting targetView;
                    using (var t = new Transaction(targetDoc, "KPLN. Создание временного вида узла"))
                    {
                        t.Start();

                        var draftingTypeId = new FilteredElementCollector(targetDoc)
                            .OfClass(typeof(ViewFamilyType))
                            .Cast<ViewFamilyType>()
                            .First(vft => vft.ViewFamily == ViewFamily.Drafting)
                            .Id;

                        targetView = ViewDrafting.Create(targetDoc, draftingTypeId);
                        targetView.Name = "_tempViewForPlugin0517";

                        t.Commit();
                    }

                    var elementsToCopy = allInView.Where(e => e.ViewSpecific && !(e is ImportInstance)).Select(e => e.Id).ToList();

                    if (elementsToCopy.Count == 0)
                    {
                        if (openedSourceHere)
                        {
                            try { sourceDoc.Close(false); } catch { }
                        }

                        _uiapp.ActiveUIDocument.ActiveView = targetView;
                        TaskDialog.Show("KPLN. Менеджер узлов", $"Создан вид \"{targetView.Name}\", но на исходном виде нет элементов для копирования.");
                        return;
                    }

                    var options = new CopyPasteOptions();
                    options.SetDuplicateTypeNamesHandler(new UseDestinationTypesHandler());

                    using (var t = new Transaction(targetDoc, "KPLN. Копирование элементов узла"))
                    {
                        t.Start();

                        ElementTransformUtils.CopyElements(sourceView, elementsToCopy, targetView, Autodesk.Revit.DB.Transform.Identity, options);

                        t.Commit();
                    }

                    bool hasElemsOnTempView = new FilteredElementCollector(targetDoc, targetView.Id).WhereElementIsNotElementType().Any(e => e.ViewSpecific);

                    using (var t = new Transaction(targetDoc, "KPLN. Обработка временного вида узла"))
                    {
                        t.Start();
                        targetDoc.Delete(targetView.Id);
                        targetView = null;
                        t.Commit();
                    }

                    if (openedSourceHere)
                    {
                        try { sourceDoc.Close(false); } catch { }
                    }

                    TaskDialog.Show("KPLN. Менеджер узлов", $"Вид с узлом ``{sourceViewName}`` скопирован в активный документ.");
                }
            }

            catch (Exception ex)
            {
                TaskDialog.Show("KPLN. Менеджер узлов", "Ошибка:\n" + ex.Message);
            }
            finally
            {
                try
                {
                    _ownerWindow?.Dispatcher.Invoke(() =>
                    {
                        _ownerWindow.Topmost = true;
                        _ownerWindow.Activate();
                    });
                }
                catch
                {
                }
                finally
                {
                    if (_ownerWindow != null && !_ownerWindow.Dispatcher.HasShutdownStarted)
                    {
                        _ownerWindow.Dispatcher.BeginInvoke(
                            new Action(() => _ownerWindow.Topmost = false),
                            System.Windows.Threading.DispatcherPriority.ApplicationIdle);
                    }
                }
            }
        }
    }


















    ////////////////////// КОПИРОВАНИЕ НА ВИД, ЛИСТ, ЛЕГЕНДУ
    internal sealed class PlaceDraftingViewOnSheetHandler : IExternalEventHandler
    {
        private UIApplication _uiapp;
        private string _sourceRvtPath;
        private string _sourceViewName;
        private Window _ownerWindow;

        private enum DwgChoice
        {
            ExportOverwriteAndUse,
            UseExisting,
            Cancel
        }

        public void Init(UIApplication uiapp, string sourceRvtPath, string sourceViewName, Window ownerWindow)
        {
            _uiapp = uiapp;
            _sourceRvtPath = sourceRvtPath;
            _sourceViewName = sourceViewName;
            _ownerWindow = ownerWindow;
        }

        public string GetName() => "KPLN. Размещение узла";

        private static readonly string _stateDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "KPLN", "МУ");
        private static readonly string _lastDwgDirFile = Path.Combine(_stateDir, "lastDirDWG.txt");

        private static string LoadLastDwgDir()
        {
            try
            {
                if (!File.Exists(_lastDwgDirFile)) return null;
                var dir = (File.ReadAllText(_lastDwgDirFile) ?? string.Empty).Trim();
                return Directory.Exists(dir) ? dir : null;
            }
            catch { return null; }
        }

        private static void SaveLastDwgDir(string dir)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(dir)) return;
                if (!Directory.Exists(dir)) return;

                Directory.CreateDirectory(_stateDir);
                File.WriteAllText(_lastDwgDirFile, dir);
            }
            catch { }
        }

        private static bool HasNonDwgViewSpecificStuff(Document donorDoc, View donorView)
        {
            try
            {
                return new FilteredElementCollector(donorDoc, donorView.Id)
                    .WhereElementIsNotElementType()
                    .Any(e =>
                        e != null &&
                        e.ViewSpecific &&
                        !(e is ImportInstance) &&
                        !(e is View) &&
                        !(e is Group) &&
                        e.Category != null);
            }
            catch
            {
                return true;
            }
        }

        private static void AbortIfDwgMixedWithGraphics(Document donorDoc, ViewDrafting donorView, bool hasDwg)
        {
            if (!hasDwg)
                return;

            bool hasNonDwgStuff = HasNonDwgViewSpecificStuff(donorDoc, donorView);
            if (!hasNonDwgStuff)
                return;

            TaskDialog.Show("Менеджер узлов",
                "Вы пытаетесь добавить узел, который состоит одновременно из DWG и встроенной графики.\n" +
                "Добавление таких узлов запрещено: обратитесь к разработчику узла для того, чтобы он поправил данный узел.");
            throw new OperationCanceledException("DWG + встроенная графика на донорском виде");
        }

        private static bool TargetCentralIsNotLocalFile(Document doc, out string centralUserVisible)
        {
            centralUserVisible = null;

            try
            {
                if (!doc.IsWorkshared) return false;

                var mp = doc.GetWorksharingCentralModelPath();
                if (mp == null) return false;

                centralUserVisible = ModelPathUtils.ConvertModelPathToUserVisiblePath(mp) ?? "";

                if (centralUserVisible.StartsWith("RSN://", StringComparison.InvariantCultureIgnoreCase) ||
                    centralUserVisible.StartsWith("RSN:", StringComparison.InvariantCultureIgnoreCase))
                    return true;

                if (string.IsNullOrWhiteSpace(centralUserVisible)) return true;
                if (!File.Exists(centralUserVisible)) return true;

                return false;
            }
            catch
            {
                return true;
            }
        }

        private static string AskUserForDwgPath(Window ownerWindow, string centralInfo, string defaultFileName)
        {
            var lastDir = LoadLastDwgDir();

            string startDir =
                (!string.IsNullOrWhiteSpace(lastDir) && Directory.Exists(lastDir)) ? lastDir :
                Directory.Exists(@"Y:\") ? @"Y:\" :
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Необходимо выбрать локальную дирректорию, чтобы сохранить DWG",
                Filter = "DWG (*.dwg)|*.dwg",
                DefaultExt = ".dwg",
                AddExtension = true,
                OverwritePrompt = false,
                FileName = defaultFileName,
                InitialDirectory = startDir,
                RestoreDirectory = true
            };

            bool? res = dlg.ShowDialog(ownerWindow);
            if (res != true || string.IsNullOrWhiteSpace(dlg.FileName))
                return null;

            var folder = Path.GetDirectoryName(dlg.FileName);
            if (!string.IsNullOrWhiteSpace(folder) && Directory.Exists(folder))
                SaveLastDwgDir(folder);

            return dlg.FileName;
        }

        private string GetDwgFilePathForTarget(Document targetDoc)
        {
            if (TargetCentralIsNotLocalFile(targetDoc, out var centralStr))
            {
                string dwgFolderFromDb = ResolveDwgFolderFromMainDb(centralStr);
                if (string.IsNullOrWhiteSpace(dwgFolderFromDb))
                    return null;

                return Path.Combine(dwgFolderFromDb, _sourceViewName + ".dwg");
            }

            string baseModelPath = GetTargetModelPath(targetDoc);
            if (string.IsNullOrWhiteSpace(baseModelPath) || !File.Exists(baseModelPath))
                return null;

            string modelDir = Path.GetDirectoryName(baseModelPath);
            if (string.IsNullOrWhiteSpace(modelDir))
                return null;

            string dwgFolder = Path.Combine(modelDir, "DWG_Менеджер узлов");
            Directory.CreateDirectory(dwgFolder);

            return Path.Combine(dwgFolder, _sourceViewName + ".dwg");
        }

        public void Execute(UIApplication app)
        {
            var uidoc = app.ActiveUIDocument;
            if (uidoc == null)
                return;

            var doc = uidoc.Document;
            if (doc == null)
                return;

            var av = uidoc.ActiveView;
            bool isLegend = av.ViewType == ViewType.Legend;
            if (!(av is ViewSheet) && !(av is ViewDrafting) && !isLegend)
            {
                TaskDialog.Show("KPLN. Менеджер узлов", "Копирование узла поддерживается только, если открыт лист, чертёжный вид или легенда.");
                return;
            }
            if (string.IsNullOrWhiteSpace(_sourceRvtPath) || string.IsNullOrWhiteSpace(_sourceViewName))
            {
                TaskDialog.Show("KPLN. Менеджер узлов", "Не заданы путь к донорской модели или имя вида-донора.");
                return;
            }

            if (av is ViewSheet sheet)
            {
                _ownerWindow.Topmost = false;
                HandleSheetCase(app, uidoc, doc, sheet);
            }
            else if (av is ViewDrafting draftingView)
            {
                _ownerWindow.Topmost = false;
                HandleDraftingCase(app, uidoc, doc, draftingView);
            }
            else if (av.ViewType == ViewType.Legend)
            {
                _ownerWindow.Topmost = false;
                HandleLegendCase(app, uidoc, doc, av);
            }

            _ownerWindow.Topmost = true;
            _ownerWindow.Topmost = false;

        }

        /// <summary>
        /// Проверка, есть ли на виде DWG.
        /// </summary>
        private bool ViewContainsDwg(Document donorDoc, ViewDrafting donorView)
        {
            var imports = new FilteredElementCollector(donorDoc, donorView.Id)
                .OfClass(typeof(ImportInstance)).Cast<ImportInstance>().ToList();

            foreach (var imp in imports)
            {
                var type = donorDoc.GetElement(imp.GetTypeId()) as ElementType;
                if (type == null) continue;

                var name = type.Name ?? string.Empty;
                if (name.EndsWith(".dwg", StringComparison.InvariantCultureIgnoreCase))
                    return true;
            }

            return false;
        }






















        private static string ResolveDwgFolderFromMainDb(string centralPath)
        {
            // centralPath должен быть именно CentralPath (например RSN://...)
            if (string.IsNullOrWhiteSpace(centralPath))
                return null;

            string mainDb = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_Loader_MainDB.db";

            string foundMainPath = null;
            int? foundSubDepId = null;
            string foundSubDepCode = null;

            try
            {
                using (var conn = new System.Data.SQLite.SQLiteConnection(
                    "Data Source=" + mainDb + ";Version=3;FailIfMissing=True;"))
                {
                    conn.Open();

                    // 1) Projects -> MainPath по префиксу RevitServerPath*
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = @"
                    SELECT MainPath, RevitServerPath, RevitServerPath2, RevitServerPath3, RevitServerPath4
                    FROM Projects
                    WHERE (RevitServerPath  IS NOT NULL AND TRIM(RevitServerPath)  <> '')
                       OR (RevitServerPath2 IS NOT NULL AND TRIM(RevitServerPath2) <> '')
                       OR (RevitServerPath3 IS NOT NULL AND TRIM(RevitServerPath3) <> '')
                       OR (RevitServerPath4 IS NOT NULL AND TRIM(RevitServerPath4) <> '');";

                        using (var r = cmd.ExecuteReader())
                        {
                            while (r.Read())
                            {
                                string mp = r.IsDBNull(0) ? null : r.GetString(0);

                                string p1 = r.IsDBNull(1) ? null : r.GetString(1);
                                string p2 = r.IsDBNull(2) ? null : r.GetString(2);
                                string p3 = r.IsDBNull(3) ? null : r.GetString(3);
                                string p4 = r.IsDBNull(4) ? null : r.GetString(4);

                                bool match =
                                    (!string.IsNullOrWhiteSpace(p1) && centralPath.StartsWith(p1.Trim(), StringComparison.InvariantCultureIgnoreCase)) ||
                                    (!string.IsNullOrWhiteSpace(p2) && centralPath.StartsWith(p2.Trim(), StringComparison.InvariantCultureIgnoreCase)) ||
                                    (!string.IsNullOrWhiteSpace(p3) && centralPath.StartsWith(p3.Trim(), StringComparison.InvariantCultureIgnoreCase)) ||
                                    (!string.IsNullOrWhiteSpace(p4) && centralPath.StartsWith(p4.Trim(), StringComparison.InvariantCultureIgnoreCase));

                                if (match)
                                {
                                    foundMainPath = mp;
                                    break;
                                }
                            }
                        }
                    }

                    if (string.IsNullOrWhiteSpace(foundMainPath))
                    {
                        TaskDialog.Show("KPLN. Менеджер узлов",
                            "В MainDB в таблице Projects не найдено соответствий RevitServerPath* и MainPath.\n" +
                            "Текущий CentralPath:\n" + centralPath + "\n" +
                            "Для решения проблемы - обратитесь в BIM-отдел");
                        return null;
                    }

                    // 2) Documents -> SubDepartmentId по CentralPath
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = @"
                    SELECT SubDepartmentId
                    FROM Documents
                    WHERE CentralPath = @p
                    LIMIT 1;";
                        cmd.Parameters.AddWithValue("@p", centralPath);

                        var obj = cmd.ExecuteScalar();
                        if (obj != null && obj != DBNull.Value && int.TryParse(obj.ToString(), out int dep))
                            foundSubDepId = dep;
                    }

                    if (!foundSubDepId.HasValue)
                    {
                        TaskDialog.Show("KPLN. Менеджер узлов",
                            "Файл с данным CentralPath не зарегистрирован в БД (таблица Documents).\n" +
                            "Текущий CentralPath:\n" + centralPath + "\n" +
                            "Для решения проблемы - обратитесь в BIM-отдел");
                        return null;
                    }

                    // 3) SubDepartments -> Code по Id
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = @"
                    SELECT Code
                    FROM SubDepartments
                    WHERE Id = @id
                    LIMIT 1;";
                        cmd.Parameters.AddWithValue("@id", foundSubDepId.Value);

                        var obj = cmd.ExecuteScalar();
                        foundSubDepCode = (obj == null || obj == DBNull.Value) ? null : obj.ToString();
                        if (!string.IsNullOrWhiteSpace(foundSubDepCode))
                            foundSubDepCode = foundSubDepCode.Trim();
                    }

                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("KPLN. Менеджер узлов", "Ошибка чтения MainDB:\n" + ex.Message);
                return null;
            }

            // приводим путь проекта к Y:\
            string mainPathDisplay = foundMainPath.Replace(@"\\stinproject.local\project\", @"Y:\");

            // отдел
            string depDisplay = !string.IsNullOrWhiteSpace(foundSubDepCode)
                ? foundSubDepCode
                : foundSubDepId.Value.ToString(CultureInfo.InvariantCulture);

            // чистим имя папки
            foreach (char ch in Path.GetInvalidFileNameChars())
                depDisplay = depDisplay.Replace(ch.ToString(), "_");

            // BIM\8.Менеджер узлов\<dep>
            string baseBim = Path.Combine(mainPathDisplay, "BIM");
            string managerDir = Path.Combine(baseBim, "8.Менеджер узлов");
            string dwgFolder = Path.Combine(managerDir, depDisplay);

            try
            {
                Directory.CreateDirectory(dwgFolder);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("KPLN. Менеджер узлов",
                    "Не удалось создать директорию для DWG:\n" + dwgFolder + "\n\n" + ex.Message);
                return null;
            }

            return dwgFolder;
        }





















        /// <summary>
        /// ЛИСТ
        /// </summary>
        private void HandleSheetCase(UIApplication app, UIDocument uidoc, Document targetDoc, ViewSheet targetSheet)
        {
            Document donorDoc = null;
            bool donorOpenedHere = false;

            try
            {
                donorDoc = app.Application.Documents.Cast<Document>()
                    .FirstOrDefault(d => !string.IsNullOrEmpty(d.PathName) && string.Equals(d.PathName, _sourceRvtPath, StringComparison.InvariantCultureIgnoreCase));

                if (donorDoc == null)
                {
                    var modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(_sourceRvtPath);
                    var openOpts = new OpenOptions
                    {
                        DetachFromCentralOption = DetachFromCentralOption.DoNotDetach
                    };
                    donorDoc = app.Application.OpenDocumentFile(modelPath, openOpts);
                    donorOpenedHere = true;
                }

                if (donorDoc == null)
                {
                    TaskDialog.Show("KPLN. Менеджер узлов. Ошибка", "Не удалось открыть модель-донора");
                    return;
                }

                var donorView = new FilteredElementCollector(donorDoc).OfClass(typeof(ViewDrafting))
                    .Cast<ViewDrafting>().FirstOrDefault(v => !v.IsTemplate && string.Equals(v.Name, _sourceViewName, StringComparison.InvariantCultureIgnoreCase));

                if (donorView == null)
                {
                    TaskDialog.Show("KPLN. Менеджер узлов", $"В файле-донора не найден чертёжный вид с именем:\n\"{_sourceViewName}\".");
                    return;
                }

                bool hasDwg = ViewContainsDwg(donorDoc, donorView);


                try
                {
                    AbortIfDwgMixedWithGraphics(donorDoc, donorView, hasDwg);
                }
                catch (OperationCanceledException)
                {
                    return;
                }

                // DWG
                if (hasDwg)
                {
                    HandleDwgBranch(donorDoc, donorView, uidoc, targetDoc, targetSheet);
                }
                // Без DWG
                else
                {
                    HandleNonDwgBranch(donorDoc, donorView, uidoc, targetDoc, targetSheet);
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("KPLN. Менеджер узлов", "Ошибка при размещении узла:\n" + ex.Message);
            }
            finally
            {
                if (donorOpenedHere && donorDoc != null && donorDoc.IsValidObject)
                {
                    try { donorDoc.Close(false); }
                    catch { }
                }
            }
        }

        /// <summary>
        /// ЛИСТ. На донорском виде есть DWG.
        /// </summary>
        private void HandleDwgBranch(Document donorDoc, ViewDrafting donorView, UIDocument uidoc, Document targetDoc, ViewSheet targetSheet)
        {
            string dwgFilePath = GetDwgFilePathForTarget(targetDoc);
            if (string.IsNullOrWhiteSpace(dwgFilePath))
            {
                TaskDialog.Show("KPLN. Менеджер узлов", "Не удалось определить путь для DWG.");
                return;
            }

            string dwgFolder = Path.GetDirectoryName(dwgFilePath);
            string dwgNameNoExt = Path.GetFileNameWithoutExtension(dwgFilePath);

            if (string.IsNullOrWhiteSpace(dwgFolder) || string.IsNullOrWhiteSpace(dwgNameNoExt))
            {
                TaskDialog.Show("KPLN. Менеджер узлов", "Некорректный путь для DWG.");
                return;
            }

            bool dwgExists = File.Exists(dwgFilePath);

            DwgChoice choice;

            if (dwgExists)
            {
                var td = new TaskDialog("KPLN. Менеджер узлов");
                td.MainInstruction = $"Найден DWG с именем \"{_sourceViewName}.dwg\" в папке:\n{dwgFolder}";
                td.MainContent = "Выберите, что сделать с DWG:";
                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1,
                    "Перезаписать DWG и использовать его");
                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2,
                    "Использовать существующий DWG из папки");
                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink3,
                    "Отмена");
                td.CommonButtons = TaskDialogCommonButtons.Close;
                td.DefaultButton = TaskDialogResult.Close;

                var res = td.Show();
                switch (res)
                {
                    case TaskDialogResult.CommandLink1:
                        choice = DwgChoice.ExportOverwriteAndUse;
                        break;
                    case TaskDialogResult.CommandLink2:
                        choice = DwgChoice.UseExisting;
                        break;
                    default:
                        choice = DwgChoice.Cancel;
                        break;
                }
            }
            else
            {
                TaskDialog.Show("Выбор точки","Укажите точку для размещения узла");
                choice = DwgChoice.ExportOverwriteAndUse;
            }

            if (choice == DwgChoice.Cancel)
                return;

            if (choice == DwgChoice.ExportOverwriteAndUse)
            {
                if (File.Exists(dwgFilePath))
                {
                    try { File.Delete(dwgFilePath); }
                    catch (Exception ex)
                    {
                        TaskDialog.Show("KPLN. Менеджер узлов", "Не удалось удалить существующий DWG:\n" + ex.Message);
                        return;
                    }
                }

                var dwgExportOptions = new DWGExportOptions();
                var viewIds = new List<ElementId> { donorView.Id };

                bool exported = donorDoc.Export(dwgFolder, _sourceViewName, viewIds, dwgExportOptions);
                if (!exported || !File.Exists(dwgFilePath))
                {
                    TaskDialog.Show("KPLN. Менеджер узлов", "Не удалось экспортировать DWG из вида-донора.");
                    return;
                }
            }
            else
            {
                if (!File.Exists(dwgFilePath))
                {
                    TaskDialog.Show("KPLN. Менеджер узлов",
                        "Файл DWG из папки не найден, хотя должен быть. Операция прервана.");
                    return;
                }
            }

            ViewDrafting draftingView = null;

            using (var t = new Transaction(targetDoc, "KPLN. Обновление вида узла (DWG)"))
            {
                t.Start();

                draftingView = new FilteredElementCollector(targetDoc)
                    .OfClass(typeof(ViewDrafting)).Cast<ViewDrafting>()
                    .FirstOrDefault(v => !v.IsTemplate && string.Equals(v.Name, _sourceViewName, StringComparison.InvariantCultureIgnoreCase));

                if (draftingView == null)
                {
                    var draftingType = new FilteredElementCollector(targetDoc)
                        .OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault(vft => vft.ViewFamily == ViewFamily.Drafting);

                    if (draftingType == null)
                    {
                        TaskDialog.Show("KPLN. Менеджер узлов", "В проекте не найден тип для чертёжных видов.");
                        t.RollBack();
                        return;
                    }

                    draftingView = ViewDrafting.Create(targetDoc, draftingType.Id);
                    draftingView.Name = _sourceViewName;
                }

                var oldImports = new FilteredElementCollector(targetDoc, draftingView.Id)
                    .OfClass(typeof(ImportInstance)).Cast<ImportInstance>().ToList();

                foreach (var imp in oldImports)
                {
                    try { targetDoc.Delete(imp.Id); }
                    catch { }
                }

                var dwgImportOptions = new DWGImportOptions
                {
                    ThisViewOnly = true
                };

                ElementId importedId;
                bool imported = targetDoc.Import(dwgFilePath, dwgImportOptions, draftingView, out importedId);
                if (!imported)
                {
                    TaskDialog.Show("KPLN. Менеджер узлов",
                        "Не удалось импортировать/залинковать DWG в вид узла.");
                    t.RollBack();
                    return;
                }

                t.Commit();
            }

            if (draftingView == null)
                return;

            PlaceViewportOnSheet(uidoc, targetDoc, targetSheet, draftingView);
        }

        /// <summary>
        /// ЛИСТ. На донорском виде нет DWG.
        /// </summary>
        private void HandleNonDwgBranch(Document donorDoc, ViewDrafting donorView, UIDocument uidoc, Document targetDoc, ViewSheet targetSheet)
        {
            var existingView = new FilteredElementCollector(targetDoc).OfClass(typeof(ViewDrafting))
                .Cast<ViewDrafting>().FirstOrDefault(v => !v.IsTemplate && string.Equals(v.Name, _sourceViewName, StringComparison.InvariantCultureIgnoreCase));

            if (existingView != null)
            {
                var td = new TaskDialog("KPLN. Менеджер узлов");
                td.MainInstruction = $"Вид узла \"{_sourceViewName}\" уже существует в текущем файле.";
                td.MainContent = "Выберите, что сделать:";
                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1,
                    "Использовать существующий вид для размещения");
                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2,
                    "Пересоздать существующий вид из файла-донора");
                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink3,
                    "Отмена");
                td.CommonButtons = TaskDialogCommonButtons.Close;
                td.DefaultButton = TaskDialogResult.Close;

                var res = td.Show();
                if (res == TaskDialogResult.CommandLink1)
                {
                    uidoc.ActiveView = targetSheet;
                    PlaceViewportOnSheet(uidoc, targetDoc, targetSheet, existingView);
                    return;
                }
                else if (res == TaskDialogResult.CommandLink2)
                {
                    using (var tDel = new Transaction(targetDoc, "KPLN. Удаление существующего вида узла"))
                    {
                        tDel.Start();

                        var vpToDelete = new FilteredElementCollector(targetDoc).OfClass(typeof(Viewport))
                            .Cast<Viewport>().Where(vp => vp.ViewId == existingView.Id).Select(vp => vp.Id).ToList();

                        if (vpToDelete.Count > 0)
                            targetDoc.Delete(vpToDelete);

                        targetDoc.Delete(existingView.Id);

                        tDel.Commit();
                    }
                }
                else
                {
                    return;
                }
            }

            var allInView = new FilteredElementCollector(donorDoc, donorView.Id).WhereElementIsNotElementType().ToList();
            var elementsToCopy = allInView.Where(e => e.ViewSpecific && !(e is ImportInstance)).Select(e => e.Id).ToList();

            if (elementsToCopy.Count == 0)
            {
                TaskDialog.Show("KPLN. Менеджер узлов", $"На исходном виде \"{donorView.Name}\" нет элементов для копирования.");
                return;
            }

            ViewDrafting targetView = null;
            bool targetViewWasCreatedHere = false;

            using (var t = new Transaction(targetDoc, "KPLN. Подготовка вида узла"))
            {
                t.Start();

                targetView = new FilteredElementCollector(targetDoc).OfClass(typeof(ViewDrafting)).Cast<ViewDrafting>()
                    .FirstOrDefault(v => !v.IsTemplate && string.Equals(v.Name, _sourceViewName, StringComparison.InvariantCultureIgnoreCase));

                if (targetView == null)
                {
                    var draftingType = new FilteredElementCollector(targetDoc).OfClass(typeof(ViewFamilyType))
                        .Cast<ViewFamilyType>().FirstOrDefault(vft => vft.ViewFamily == ViewFamily.Drafting);

                    if (draftingType == null)
                    {
                        TaskDialog.Show("KPLN. Менеджер узлов", "В проекте не найден тип для чертёжных видов.");
                        t.RollBack();
                        return;
                    }

                    targetView = ViewDrafting.Create(targetDoc, draftingType.Id);
                    targetView.Name = _sourceViewName;
                    targetViewWasCreatedHere = true;
                }
                else
                {
                    var toDelete = new FilteredElementCollector(targetDoc, targetView.Id)
                        .WhereElementIsNotElementType().Where(e => e.ViewSpecific && !(e is ImportInstance))
                        .Select(e => e.Id).ToList();

                    if (toDelete.Count > 0)
                        targetDoc.Delete(toDelete);
                }

                t.Commit();
            }

            if (targetView == null)
                return;
            using (var t2 = new Transaction(targetDoc, "KPLN. Копирование элементов узла"))
            {
                t2.Start();

                var options = new CopyPasteOptions();
                options.SetDuplicateTypeNamesHandler(new UseDestinationTypesHandler());

                var draftingIdsBefore = new FilteredElementCollector(targetDoc)
                    .OfClass(typeof(ViewDrafting)).ToElementIds().ToList();

                try
                {
                    ElementTransformUtils.CopyElements(donorView, elementsToCopy, targetView, Autodesk.Revit.DB.Transform.Identity, options);
                }
                catch (Exception ex)
                {
                    t2.RollBack();
                    TaskDialog.Show("KPLN. Менеджер узлов", "Ошибка при копировании элементов узла:\n" + ex.Message);
                    return;
                }

                var draftingIdsAfter = new FilteredElementCollector(targetDoc).OfClass(typeof(ViewDrafting)).ToElementIds().ToList();

                var newIds = draftingIdsAfter.Where(id => !draftingIdsBefore.Contains(id)).ToList();

                if (newIds.Count > 0)
                {
                    var realView = targetDoc.GetElement(newIds[0]) as ViewDrafting;
                    if (realView != null)
                    {
                        if (targetViewWasCreatedHere && targetView.Id != realView.Id)
                        {
                            try { targetDoc.Delete(targetView.Id); } catch { }
                        }

                        targetView = realView;

                        if (!string.Equals(targetView.Name, _sourceViewName, StringComparison.InvariantCultureIgnoreCase))
                        {
                            try { targetView.Name = _sourceViewName; } catch { }
                        }
                    }
                }

                t2.Commit();
            }

            uidoc.ActiveView = targetSheet;
            TaskDialog.Show("KPLN. Менеджер узлов", "Сейчас вас перекинет на ранее ативный вид, и вам нужно будет указать точку для размещения узла на листе.");
            PlaceViewportOnSheet(uidoc, targetDoc, targetSheet, targetView);
        }

        /// <summary>
        /// Выбор точки на листе и создание/перемещение viewport.
        /// </summary>
        private void PlaceViewportOnSheet(UIDocument uidoc, Document targetDoc, ViewSheet targetSheet, View view)
        {
            uidoc.ActiveView = targetSheet;

            XYZ point;
            try
            {
                point = uidoc.Selection.PickPoint("Выберите точку для размещения узла на листе");
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return;
            }

            using (var t = new Transaction(targetDoc, "KPLN. Размещение узла на листе"))
            {
                t.Start();

                var existingVp = new FilteredElementCollector(targetDoc)
                    .OfClass(typeof(Viewport)).Cast<Viewport>().FirstOrDefault(vp => vp.ViewId == view.Id);

                if (existingVp != null)
                {
                    if (existingVp.SheetId == targetSheet.Id)
                    {
                        XYZ currentCenter = existingVp.GetBoxCenter();
                        XYZ delta = point - currentCenter;
                        ElementTransformUtils.MoveElement(targetDoc, existingVp.Id, delta);
                    }
                    else
                    {
                        var otherSheet = targetDoc.GetElement(existingVp.SheetId) as ViewSheet;
                        string sheetName = otherSheet != null ? otherSheet.SheetNumber + " " + otherSheet.Name : "<неизвестно>";

                        TaskDialog.Show("KPLN. Менеджер узлов", $"Вид \"{view.Name}\" уже размещён на листе:\n{sheetName}\n\n" +
                            "Один и тот же вид нельзя добавить на несколько листов. Создайте дубликат вида, если нужно разместить узел на другом листе.");

                    }

                    t.Commit();
                    return;
                }

                Viewport.Create(targetDoc, targetSheet.Id, view.Id, point);

                var titleParam = view.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION);
                if (titleParam != null && !titleParam.IsReadOnly)
                {
                    titleParam.Set("\u200B");
                }

                t.Commit();
            }
        }

        /// <summary>
        /// ЧЕРТЁЖНЫЙ ВИД
        /// </summary>
        private void HandleDraftingCase(UIApplication app, UIDocument uidoc, Document targetDoc, ViewDrafting targetDraftingView)
        {
            Document donorDoc = null;
            bool donorOpenedHere = false;

            try
            {
                donorDoc = app.Application.Documents.Cast<Document>()
                    .FirstOrDefault(d => !string.IsNullOrEmpty(d.PathName) && string.Equals(d.PathName, _sourceRvtPath, StringComparison.InvariantCultureIgnoreCase));

                if (donorDoc == null)
                {
                    var modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(_sourceRvtPath);
                    var openOpts = new OpenOptions
                    {
                        DetachFromCentralOption = DetachFromCentralOption.DoNotDetach
                    };
                    donorDoc = app.Application.OpenDocumentFile(modelPath, openOpts);
                    donorOpenedHere = true;
                }

                if (donorDoc == null)
                {
                    TaskDialog.Show("KPLN. Менеджер узлов", $"Не удалось открыть модель-донора:\n{_sourceRvtPath}");
                    return;
                }

                var donorView = new FilteredElementCollector(donorDoc).OfClass(typeof(ViewDrafting)).Cast<ViewDrafting>()
                    .FirstOrDefault(v => !v.IsTemplate && string.Equals(v.Name, _sourceViewName, StringComparison.InvariantCultureIgnoreCase));

                if (donorView == null)
                {
                    TaskDialog.Show("KPLN. Менеджер узлов", $"В файле-донора\n{_sourceRvtPath}\nне найден чертёжный вид с именем:\n\"{_sourceViewName}\".");
                    return;
                }

                bool hasDwg = ViewContainsDwg(donorDoc, donorView);

                try
                {
                    AbortIfDwgMixedWithGraphics(donorDoc, donorView, hasDwg);
                }
                catch (OperationCanceledException)
                {
                    return;
                }

                // DWG
                if (hasDwg)
                {
                    XYZ pickPoint;
                    try
                    {
                        TaskDialog.Show("Выбор точки", "Укажите точку для размещения узла");
                        pickPoint = uidoc.Selection.PickPoint("Выберите точку для размещения DWG");
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                    {
                        return;
                    }

                    HandleDwgOnDrafting(donorDoc, donorView, targetDoc, targetDraftingView, pickPoint);
                }
                // Без DWG
                else
                {
                    HandleNonDwgOnDrafting(uidoc, donorDoc, donorView, targetDoc, targetDraftingView);
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("KPLN. Менеджер узлов", "Ошибка при размещении узла на чертёжном виде:\n" + ex.Message);
            }
            finally
            {
                if (donorOpenedHere && donorDoc != null && donorDoc.IsValidObject)
                {
                    try { donorDoc.Close(false); } catch { }
                }
            }
        }

        /// <summary>
        /// ЧЕРТЁЖНЫЙ ВИД. На донорском виде есть DWG.
        /// </summary>
        private void HandleDwgOnDrafting(Document donorDoc, ViewDrafting donorView, Document targetDoc, ViewDrafting targetDraftingView, XYZ pickPoint)
        {
            string dwgFilePath = GetDwgFilePathForTarget(targetDoc);
            if (string.IsNullOrWhiteSpace(dwgFilePath))
            {
                TaskDialog.Show("KPLN. Менеджер узлов", "Не удалось определить путь для DWG.");
                return;
            }

            string dwgFolder = Path.GetDirectoryName(dwgFilePath);
            string dwgNameNoExt = Path.GetFileNameWithoutExtension(dwgFilePath);
            bool dwgExists = File.Exists(dwgFilePath);

            DwgChoice choice;

            if (dwgExists)
            {
                var td = new TaskDialog("KPLN. Менеджер узлов");
                td.MainInstruction = $"Найден DWG \"{_sourceViewName}.dwg\" в папке:\n{dwgFolder}";
                td.MainContent = "Выберите источник DWG:";
                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1,
                    "Перезаписать DWG текущим и использовать его");
                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2,
                    "Использовать существующий DWG из папки");
                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink3,
                    "Отмена");
                td.CommonButtons = TaskDialogCommonButtons.Close;
                td.DefaultButton = TaskDialogResult.Close;

                var res = td.Show();
                switch (res)
                {
                    case TaskDialogResult.CommandLink1:
                        choice = DwgChoice.ExportOverwriteAndUse;
                        break;
                    case TaskDialogResult.CommandLink2:
                        choice = DwgChoice.UseExisting;
                        break;
                    default:
                        choice = DwgChoice.Cancel;
                        break;
                }
            }
            else
            {               
                choice = DwgChoice.ExportOverwriteAndUse;
            }

            if (choice == DwgChoice.Cancel)
                return;

            if (choice == DwgChoice.ExportOverwriteAndUse)
            {
                if (File.Exists(dwgFilePath))
                {
                    try { File.Delete(dwgFilePath); } catch { }
                }

                var dwgExportOptions = new DWGExportOptions();
                var viewIds = new List<ElementId> { donorView.Id };

                bool exported = donorDoc.Export(dwgFolder, _sourceViewName, viewIds, dwgExportOptions);
                if (!exported || !File.Exists(dwgFilePath))
                {
                    TaskDialog.Show("KPLN. Менеджер узлов", "Не удалось экспортировать DWG из вида-донора.");
                    return;
                }
            }
            else
            {
                if (!File.Exists(dwgFilePath))
                {
                    TaskDialog.Show("KPLN. Менеджер узлов", "Файл DWG из папки не найден, хотя должен быть. Операция прервана.");
                    return;
                }
            }

            using (var t = new Transaction(targetDoc, "KPLN. Обновление DWG на чертёжном виде"))
            {
                t.Start();

                var dwgImportOptions = new DWGImportOptions
                {
                    ThisViewOnly = true,
                    Placement = ImportPlacement.Origin
                };

                ElementId importedId;
                bool imported = targetDoc.Import(dwgFilePath, dwgImportOptions, targetDraftingView, out importedId);
                if (!imported)
                {
                    TaskDialog.Show("KPLN. Менеджер узлов", "Не удалось импортировать/залинковать DWG на чертёжный вид.");
                    t.RollBack();
                    return;
                }

                var dwgInstance = targetDoc.GetElement(importedId) as ImportInstance;
                if (dwgInstance != null)
                {
                    if (dwgInstance.Pinned)
                        dwgInstance.Pinned = false;

                    var bb = dwgInstance.get_BoundingBox(targetDraftingView);
                    if (bb != null)
                    {
                        XYZ center = (bb.Min + bb.Max) * 0.5;
                        XYZ delta = pickPoint - center;
                        ElementTransformUtils.MoveElement(targetDoc, importedId, delta);
                    }
                }

                t.Commit();
            }
        }

        /// <summary>
        /// ЧЕРТЁЖНЫЙ ВИД. На донорском виде нет DWG.
        /// </summary>
        private void HandleNonDwgOnDrafting(UIDocument uidoc, Document donorDoc, ViewDrafting donorView, Document targetDoc, ViewDrafting targetDraftingView)
        {

            var allInView = new FilteredElementCollector(donorDoc, donorView.Id).WhereElementIsNotElementType().ToList();

            var elementsToCopy = allInView
                .Where(e => e.ViewSpecific && !(e is ImportInstance)).Select(e => e.Id).ToList();

            if (allInView.Count == 0)
            {
                TaskDialog.Show("KPLN. Менеджер узлов", $"На исходном виде \"{donorView.Name}\" нет элементов для копирования.");
                return;
            }

            ElementId tempViewId = null;
            List<ElementId> draftingBefore;
            List<ElementId> draftingAfter;

            using (var t = new Transaction(targetDoc, "KPLN. Копирование узла на активный чертёжный вид"))
            {
                t.Start();

                var options = new CopyPasteOptions();
                options.SetDuplicateTypeNamesHandler(new UseDestinationTypesHandler());

                draftingBefore = new FilteredElementCollector(targetDoc).OfClass(typeof(ViewDrafting)).ToElementIds().ToList();

                try
                {
                    ElementTransformUtils.CopyElements(donorView, elementsToCopy, targetDraftingView, Autodesk.Revit.DB.Transform.Identity, options);
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    TaskDialog.Show("KPLN. Менеджер узлов", "Ошибка при копировании элементов узла на чертёжный вид:\n" + ex.Message);
                    return;
                }

                draftingAfter = new FilteredElementCollector(targetDoc).OfClass(typeof(ViewDrafting)).ToElementIds().ToList();

                t.Commit();
            }

            var newIds = draftingAfter.Where(id => !draftingBefore.Contains(id)).ToList();
            if (newIds.Count > 0) tempViewId = newIds[0];

            var tempView = targetDoc.GetElement(tempViewId) as ViewDrafting;
            if (tempView == null)
            {
                TaskDialog.Show("KPLN. Менеджер узлов", "Не удалось получить временный чертёжный вид. Операция прервана.");
                return;
            }

            TaskDialog.Show("KPLN. Менеджер узлов", "Сейчас вид будет переключен на раннее активный. После этого выбирите точку, куда будет вставлены элементы чертёжного вида.");
            uidoc.ActiveView = targetDraftingView;

            XYZ pickPoint;
            try
            {
                pickPoint = uidoc.Selection.PickPoint("Выберите точку вставки узла на чертёжном виде");
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                using (var tDel = new Transaction(targetDoc, "KPLN. Удаление временного вида узла"))
                {
                    tDel.Start();
                    try { targetDoc.Delete(tempViewId); } catch { }
                    tDel.Commit();
                }
                return;
            }

            // Жёсткая сортировка
            var elemsOnTemp = new FilteredElementCollector(targetDoc, tempView.Id)
                .WhereElementIsNotElementType()
                .Where(e => e.ViewSpecific && !(e is ImportInstance) && !(e is View) && !(e is Group) && e.Category != null && e.OwnerViewId == tempView.Id)
                .Select(e => e.Id).ToList();

            if (elemsOnTemp.Count == 0)
            {
                TaskDialog.Show("KPLN. Менеджер узлов", "Во временном виде не найдено элементов для копирования.");
                using (var tDel = new Transaction(targetDoc, "KPLN. Удаление временного вида узла"))
                {
                    tDel.Start();
                    try { targetDoc.Delete(tempViewId); } catch { }
                    tDel.Commit();
                }
                return;
            }

            XYZ center;
            {
                BoundingBoxXYZ bb = null;

                foreach (var id in elemsOnTemp)
                {
                    var el = targetDoc.GetElement(id);
                    if (el == null) continue;

                    var elBb = el.get_BoundingBox(tempView);
                    if (elBb == null) continue;

                    if (bb == null)
                    {
                        bb = new BoundingBoxXYZ
                        {
                            Min = elBb.Min,
                            Max = elBb.Max
                        };
                    }
                    else
                    {
                        bb.Min = new XYZ(
                            Math.Min(bb.Min.X, elBb.Min.X),
                            Math.Min(bb.Min.Y, elBb.Min.Y),
                            Math.Min(bb.Min.Z, elBb.Min.Z));
                        bb.Max = new XYZ(
                            Math.Max(bb.Max.X, elBb.Max.X),
                            Math.Max(bb.Max.Y, elBb.Max.Y),
                            Math.Max(bb.Max.Z, elBb.Max.Z));
                    }
                }

                if (bb == null)
                {
                    TaskDialog.Show("KPLN. Менеджер узлов", "Не удалось определить границы узла на временном виде.");
                    using (var tDel = new Transaction(targetDoc, "KPLN. Удаление временного вида узла"))
                    {
                        tDel.Start();
                        try { targetDoc.Delete(tempViewId); } catch { }
                        tDel.Commit();
                    }
                    return;
                }

                center = (bb.Min + bb.Max) * 0.5;
            }

            using (var t2 = new Transaction(targetDoc, "KPLN. Вставка узла на чертёжный вид"))
            {
                t2.Start();

                var options2 = new CopyPasteOptions();
                options2.SetDuplicateTypeNamesHandler(new UseDestinationTypesHandler());

                var delta = pickPoint - center;
                var transform = Autodesk.Revit.DB.Transform.CreateTranslation(delta);

                try
                {
                    ElementTransformUtils.CopyElements(tempView, elemsOnTemp, targetDraftingView, transform, options2);
                }
                catch (Exception ex)
                {
                    t2.RollBack();
                    TaskDialog.Show("KPLN. Менеджер узлов", "Ошибка при вставке узла на чертёжный вид:\n" + ex.Message);
                    return;
                }

                t2.Commit();
            }

            using (var t3 = new Transaction(targetDoc, "KPLN. Удаление временного вида узла"))
            {
                t3.Start();
                try { targetDoc.Delete(tempViewId); } catch { }
                t3.Commit();
            }
        }



















        /// <summary>
        /// ЛЕГЕНДЫ
        /// </summary>
        private void HandleLegendCase(UIApplication app, UIDocument uidoc, Document targetDoc, View targetLegendView)
        {
            Document donorDoc = null;
            bool donorOpenedHere = false;

            try
            {
                if (targetLegendView == null || targetLegendView.ViewType != ViewType.Legend)
                {
                    TaskDialog.Show("KPLN. Менеджер узлов", "Целевой вид не является легендой.");
                    return;
                }

                donorDoc = app.Application.Documents.Cast<Document>()
                    .FirstOrDefault(d => !string.IsNullOrEmpty(d.PathName) &&
                                         string.Equals(d.PathName, _sourceRvtPath, StringComparison.InvariantCultureIgnoreCase));

                if (donorDoc == null)
                {
                    var modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(_sourceRvtPath);
                    var openOpts = new OpenOptions { DetachFromCentralOption = DetachFromCentralOption.DoNotDetach };
                    donorDoc = app.Application.OpenDocumentFile(modelPath, openOpts);
                    donorOpenedHere = true;
                }

                if (donorDoc == null)
                {
                    TaskDialog.Show("KPLN. Менеджер узлов", "Не удалось открыть модель-донора.");
                    return;
                }

                var donorView = new FilteredElementCollector(donorDoc)
                    .OfClass(typeof(ViewDrafting)).Cast<ViewDrafting>()
                    .FirstOrDefault(v => !v.IsTemplate &&
                                         string.Equals(v.Name, _sourceViewName, StringComparison.InvariantCultureIgnoreCase));

                if (donorView == null)
                {
                    TaskDialog.Show("KPLN. Менеджер узлов",
                        $"В файле-донора\n{_sourceRvtPath}\nне найден чертёжный вид:\n\"{_sourceViewName}\".");
                    return;
                }

                bool hasDwg = ViewContainsDwg(donorDoc, donorView);

                try
                {
                    AbortIfDwgMixedWithGraphics(donorDoc, donorView, hasDwg);
                }
                catch (OperationCanceledException)
                {
                    return;
                }

                uidoc.ActiveView = targetLegendView;

                XYZ pickPoint;
                try
                {
                    TaskDialog.Show("Менеджер узлов", "Выберите точку для размещения узла");
                    pickPoint = uidoc.Selection.PickPoint("Выберите точку для размещения узла на легенде");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return;
                }

                if (hasDwg)
                {
                    HandleDwgOnLegend(donorDoc, donorView, targetDoc, targetLegendView, pickPoint);
                    return;
                }

                var allInView = new FilteredElementCollector(donorDoc, donorView.Id)
                    .WhereElementIsNotElementType()
                    .ToList();

                var elementsToCopy = allInView
                    .Where(e => e.ViewSpecific && !(e is ImportInstance) && !(e is View) && !(e is Group) && e.Category != null)
                    .Select(e => e.Id)
                    .ToList();

                if (elementsToCopy.Count == 0)
                {
                    TaskDialog.Show("KPLN. Менеджер узлов", $"На донорском виде \"{donorView.Name}\" нет элементов для копирования.");
                    return;
                }

                XYZ donorCenter = GetElementsCenter(donorDoc, donorView, elementsToCopy);
                if (donorCenter == null)
                {
                    TaskDialog.Show("KPLN. Менеджер узлов", "Не удалось определить границы узла на донорском виде.");
                    return;
                }

                var options = new CopyPasteOptions();
                options.SetDuplicateTypeNamesHandler(new UseDestinationTypesHandler());

                using (var t = new Transaction(targetDoc, "KPLN. Вставка узла на легенду"))
                {
                    t.Start();

                    var delta = pickPoint - donorCenter;
                    var tr = Autodesk.Revit.DB.Transform.CreateTranslation(delta);

                    ElementTransformUtils.CopyElements(donorView, elementsToCopy, targetLegendView, tr, options);

                    t.Commit();
                }

                TaskDialog.Show("KPLN. Менеджер узлов", $"Узел \"{_sourceViewName}\" вставлен на легенду.");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("KPLN. Менеджер узлов", "Ошибка при вставке на легенду:\n" + ex.Message);
            }
            finally
            {
                if (donorOpenedHere && donorDoc != null && donorDoc.IsValidObject)
                {
                    try { donorDoc.Close(false); } catch { }
                }
            }
        }

        private void HandleDwgOnLegend(Document donorDoc, ViewDrafting donorView, Document targetDoc, View targetLegendView, XYZ pickPoint)
        {
            string dwgFilePath = GetDwgFilePathForTarget(targetDoc);
            if (string.IsNullOrWhiteSpace(dwgFilePath))
            {
                TaskDialog.Show("KPLN. Менеджер узлов", "Не удалось определить путь для DWG.");
                return;
            }

            string dwgFolder = Path.GetDirectoryName(dwgFilePath);
            string dwgNameNoExt = Path.GetFileNameWithoutExtension(dwgFilePath);
            string expectedDwgFileName = Path.GetFileName(dwgFilePath);

            if (string.IsNullOrWhiteSpace(dwgFolder) || string.IsNullOrWhiteSpace(dwgNameNoExt))
            {
                TaskDialog.Show("KPLN. Менеджер узлов", "Некорректный путь для DWG.");
                return;
            }

            bool dwgExists = File.Exists(dwgFilePath);

            DwgChoice choice;

            if (dwgExists)
            {
                var td = new TaskDialog("KPLN. Менеджер узлов");
                td.MainInstruction = $"Найден DWG \"{_sourceViewName}.dwg\" в папке:\n{dwgFolder}";
                td.MainContent = "Выберите источник DWG:";
                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Перезаписать DWG текущим и использовать его");
                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Использовать существующий DWG из папки");
                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Отмена");
                td.CommonButtons = TaskDialogCommonButtons.Close;

                var res = td.Show();
                switch (res)
                {
                    case TaskDialogResult.CommandLink1: choice = DwgChoice.ExportOverwriteAndUse; break;
                    case TaskDialogResult.CommandLink2: choice = DwgChoice.UseExisting; break;
                    default: choice = DwgChoice.Cancel; break;
                }
            }
            else
            {
                choice = DwgChoice.ExportOverwriteAndUse;
            }

            if (choice == DwgChoice.Cancel)
                return;

            if (choice == DwgChoice.ExportOverwriteAndUse)
            {
                if (File.Exists(dwgFilePath))
                {
                    try { File.Delete(dwgFilePath); } catch { }
                }

                var dwgExportOptions = new DWGExportOptions();
                var viewIds = new List<ElementId> { donorView.Id };

                bool exported = donorDoc.Export(dwgFolder, _sourceViewName, viewIds, dwgExportOptions);
                if (!exported || !File.Exists(dwgFilePath))
                {
                    TaskDialog.Show("KPLN. Менеджер узлов", "Не удалось экспортировать DWG из вида-донора.");
                    return;
                }
            }
            else
            {
                if (!File.Exists(dwgFilePath))
                {
                    TaskDialog.Show("KPLN. Менеджер узлов", "Файл DWG из папки не найден. Операция прервана.");
                    return;
                }
            }

            using (var t = new Transaction(targetDoc, "KPLN. Обновление DWG на легенде"))
            {
                t.Start();

                var dwgImportOptions = new DWGImportOptions
                {
                    ThisViewOnly = true,
                    Placement = ImportPlacement.Origin
                };

                ElementId importedId;
                bool imported = targetDoc.Import(dwgFilePath, dwgImportOptions, targetLegendView, out importedId);
                if (!imported)
                {
                    TaskDialog.Show("KPLN. Менеджер узлов", "Не удалось импортировать/залинковать DWG на легенду.");
                    t.RollBack();
                    return;
                }

                var dwgInstance = targetDoc.GetElement(importedId) as ImportInstance;
                if (dwgInstance != null)
                {
                    if (dwgInstance.Pinned) dwgInstance.Pinned = false;

                    var bb = dwgInstance.get_BoundingBox(targetLegendView);
                    if (bb != null)
                    {
                        XYZ center = (bb.Min + bb.Max) * 0.5;
                        XYZ delta = pickPoint - center;
                        ElementTransformUtils.MoveElement(targetDoc, importedId, delta);
                    }
                }

                t.Commit();
            }

            TaskDialog.Show("KPLN. Менеджер узлов", $"DWG с вида \"{_sourceViewName}\" импортирован на легенду.");
        }

        private static string GetTargetModelPath(Document targetDoc)
        {
            try
            {
                if (targetDoc.IsWorkshared)
                {
                    var centralMp = targetDoc.GetWorksharingCentralModelPath();
                    if (centralMp != null)
                        return ModelPathUtils.ConvertModelPathToUserVisiblePath(centralMp);
                }
            }
            catch { }

            return targetDoc.PathName;
        }

        private static XYZ GetElementsCenter(Document doc, View view, IList<ElementId> ids)
        {
            BoundingBoxXYZ bb = null;

            foreach (var id in ids)
            {
                var el = doc.GetElement(id);
                if (el == null) continue;

                var elBb = el.get_BoundingBox(view);
                if (elBb == null) continue;

                if (bb == null) bb = new BoundingBoxXYZ { Min = elBb.Min, Max = elBb.Max };
                else
                {
                    bb.Min = new XYZ(
                        Math.Min(bb.Min.X, elBb.Min.X),
                        Math.Min(bb.Min.Y, elBb.Min.Y),
                        Math.Min(bb.Min.Z, elBb.Min.Z));
                    bb.Max = new XYZ(
                        Math.Max(bb.Max.X, elBb.Max.X),
                        Math.Max(bb.Max.Y, elBb.Max.Y),
                        Math.Max(bb.Max.Z, elBb.Max.Z));
                }
            }

            return bb == null ? null : (bb.Min + bb.Max) * 0.5;
        }
    }






























    // НОЖНИЦЫ
    public class ScreenCaptureWindow : Window
    {
        private readonly int _maxWidth;
        private readonly int _maxHeight;

        private System.Windows.Controls.Canvas _canvas;
        private System.Windows.Shapes.Rectangle _selectionRectShape;
        private System.Windows.Point _startPoint;
        private bool _isDragging;

        public byte[] CapturedBytes { get; private set; }

        public ScreenCaptureWindow(int maxWidth, int maxHeight)
        {
            _maxWidth = maxWidth;
            _maxHeight = maxHeight;

            InitWindow();
        }

        private void InitWindow()
        {
            this.WindowStyle = WindowStyle.None;
            this.AllowsTransparency = true;
            this.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(10, 0, 0, 0));
            this.ResizeMode = ResizeMode.NoResize;
            this.ShowInTaskbar = false;

            this.Left = SystemParameters.VirtualScreenLeft;
            this.Top = SystemParameters.VirtualScreenTop;
            this.Width = SystemParameters.VirtualScreenWidth;
            this.Height = SystemParameters.VirtualScreenHeight;

            this.Cursor = Cursors.Cross;

            _canvas = new System.Windows.Controls.Canvas();
            this.Content = _canvas;

            _selectionRectShape = new System.Windows.Shapes.Rectangle
            {
                Stroke = System.Windows.Media.Brushes.Red,
                StrokeThickness = 1,
                Fill = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromArgb(50, 255, 255, 255)),
                Visibility = System.Windows.Visibility.Collapsed
            };
            _canvas.Children.Add(_selectionRectShape);

            this.MouseDown += OnMouseDown;
            this.MouseMove += OnMouseMove;
            this.MouseUp += OnMouseUp;
            this.KeyDown += OnKeyDown;
        }

        private void OnKeyDown(object sender, KeyEventArgs e)
        {

            if (e.Key == Key.Escape)
            {
                this.DialogResult = false;
                this.Close();
            }
        }

        private void OnMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed)
                return;

            _isDragging = true;
            _startPoint = e.GetPosition(this);

            System.Windows.Controls.Canvas.SetLeft(_selectionRectShape, _startPoint.X);
            System.Windows.Controls.Canvas.SetTop(_selectionRectShape, _startPoint.Y);
            _selectionRectShape.Width = 0;
            _selectionRectShape.Height = 0;
            _selectionRectShape.Visibility = System.Windows.Visibility.Visible;
        }

        private void OnMouseMove(object sender, MouseEventArgs e)
        {
            if (!_isDragging)
                return;

            System.Windows.Point pos = e.GetPosition(this);

            double x = Math.Min(pos.X, _startPoint.X);
            double y = Math.Min(pos.Y, _startPoint.Y);
            double w = Math.Abs(pos.X - _startPoint.X);
            double h = Math.Abs(pos.Y - _startPoint.Y);

            System.Windows.Controls.Canvas.SetLeft(_selectionRectShape, x);
            System.Windows.Controls.Canvas.SetTop(_selectionRectShape, y);
            _selectionRectShape.Width = w;
            _selectionRectShape.Height = h;
        }

        private void OnMouseUp(object sender, MouseButtonEventArgs e)
        {
            if (!_isDragging)
                return;

            _isDragging = false;

            double x = System.Windows.Controls.Canvas.GetLeft(_selectionRectShape);
            double y = System.Windows.Controls.Canvas.GetTop(_selectionRectShape);
            double w = _selectionRectShape.Width;
            double h = _selectionRectShape.Height;

            if (w < 5 || h < 5)
            {
                this.DialogResult = false;
                this.Close();
                return;
            }

            var dpi = System.Windows.Media.VisualTreeHelper.GetDpi(this);
            int leftPx = (int)Math.Round((this.Left + x) * dpi.DpiScaleX);
            int topPx = (int)Math.Round((this.Top + y) * dpi.DpiScaleY);
            int widthPx = (int)Math.Round(w * dpi.DpiScaleX);
            int heightPx = (int)Math.Round(h * dpi.DpiScaleY);

            try
            {
                CapturedBytes = CaptureAndScale(leftPx, topPx, widthPx, heightPx,
                    _maxWidth, _maxHeight);
                this.DialogResult = CapturedBytes != null;
            }
            catch
            {
                this.DialogResult = false;
            }
            finally
            {
                this.Close();
            }
        }

        private static byte[] CaptureAndScale(
            int left, int top, int width, int height,
            int maxWidth, int maxHeight)
        {
            using (var bmp = new System.Drawing.Bitmap(width, height))
            {
                using (var g = System.Drawing.Graphics.FromImage(bmp))
                {
                    g.CopyFromScreen(left, top, 0, 0,
                        new System.Drawing.Size(width, height));
                }

                double scaleW = (double)maxWidth / width;
                double scaleH = (double)maxHeight / height;
                double scale = Math.Min(Math.Min(scaleW, scaleH), 1.0);

                int newW = (int)Math.Round(width * scale);
                int newH = (int)Math.Round(height * scale);

                System.Drawing.Bitmap resultBmp = bmp;

                if (scale < 1.0)
                {
                    resultBmp = new System.Drawing.Bitmap(newW, newH);
                    using (var g2 = System.Drawing.Graphics.FromImage(resultBmp))
                    {
                        g2.InterpolationMode =
                            System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        g2.SmoothingMode =
                            System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                        g2.PixelOffsetMode =
                            System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                        g2.CompositingQuality =
                            System.Drawing.Drawing2D.CompositingQuality.HighQuality;

                        g2.Clear(System.Drawing.Color.White);
                        g2.DrawImage(bmp, 0, 0, newW, newH);
                    }
                }

                try
                {
                    using (var ms = new MemoryStream())
                    {
                        resultBmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                        return ms.ToArray();
                    }
                }
                finally
                {
                    if (!object.ReferenceEquals(resultBmp, bmp))
                        resultBmp.Dispose();
                }
            }
        }
    }
}
