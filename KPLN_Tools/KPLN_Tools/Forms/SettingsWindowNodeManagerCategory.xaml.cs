using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.SQLite;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace KPLN_Tools.Forms
{
    public partial class SettingsWindowNodeManagerCategory : Window
    {
        private const string DbPath = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_NodeManager.db";

        private ObservableCollection<CategoryNode> _root = new ObservableCollection<CategoryNode>();
        private readonly Dictionary<string, string> _jsonByName = new Dictionary<string, string>();
        private string _currentName;

        public SettingsWindowNodeManagerCategory()
        {
            InitializeComponent();
            LoadNames();
        }

        private SQLiteConnection OpenConn()
        {
            var cs = string.Format("Data Source={0};Version=3;", DbPath);
            var conn = new SQLiteConnection(cs);
            conn.Open();
            return conn;
        }

        private void LoadNames()
        {
            try
            {
                using (var conn = OpenConn())
                using (var cmd = new SQLiteCommand("SELECT NAME, SUBCAT_JSON FROM nodeCategory;", conn))
                using (var rdr = cmd.ExecuteReader())
                {
                    var names = new List<string>();
                    _jsonByName.Clear();

                    while (rdr.Read())
                    {
                        var name = rdr["NAME"] == null ? null : rdr["NAME"].ToString();
                        var json = rdr["SUBCAT_JSON"] == null ? "[]" : rdr["SUBCAT_JSON"].ToString();

                        if (!string.IsNullOrWhiteSpace(name))
                        {
                            names.Add(name);
                            _jsonByName[name] = string.IsNullOrWhiteSpace(json) ? "[]" : json;
                        }
                    }

                    CbNames.ItemsSource = names;
                    if (names.Count > 0)
                        CbNames.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка загрузки имён из БД:\n" + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadTreeForName(string name)
        {
            _root.Clear();
            _currentName = name;

            string json;
            if (string.IsNullOrEmpty(name) || !_jsonByName.TryGetValue(name, out json))
                json = "[]";

            try
            {
                var token = string.IsNullOrWhiteSpace(json) ? new JArray() : JToken.Parse(json);

                if (token is JArray array)
                {
                    foreach (var el in array)
                    {
                        var node = ParseNode(el, null);
                        if (node != null)
                            _root.Add(node);
                    }
                }

                Tree.ItemsSource = _root;
                ExpandCollapse(Tree, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка парсинга JSON:\n" + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                _root.Clear();
                Tree.ItemsSource = _root;
            }
        }

        private void SaveCurrentToDb()
        {
            if (string.IsNullOrEmpty(_currentName))
            {
                MessageBox.Show("Не выбрана категория.", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            foreach (var n in EnumerateAll(_root))
            {
                if (n.Depth > 3)
                {
                    MessageBox.Show($"Превышена допустимая глубина вложенности (до 3): {n.Title}",
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            }

            var arr = new JArray();
            foreach (var n in _root)
                arr.Add(ToJObject(n));

            var settings = new JsonSerializerSettings
            {
                NullValueHandling = NullValueHandling.Ignore,
                Formatting = Formatting.Indented
            };
            var json = JsonConvert.SerializeObject(arr, settings);

            try
            {
                using (var conn = OpenConn())
                using (var cmd = new SQLiteCommand(
                    "UPDATE nodeCategory SET SUBCAT_JSON = @json WHERE NAME = @name;", conn))
                {
                    cmd.Parameters.AddWithValue("@json", json);
                    cmd.Parameters.AddWithValue("@name", _currentName);
                    var rows = cmd.ExecuteNonQuery();

                    if (rows == 0)
                    {
                        MessageBox.Show("Строка с таким именем категории не найдена.", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    else
                    {
                        _jsonByName[_currentName] = json;
                        MessageBox.Show("Изменения сохранены", "Готово", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка записи в БД:\n" + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private CategoryNode ParseNode(JToken token, CategoryNode parent)
        {
            if (token == null || token.Type != JTokenType.Object)
                return null;

            var obj = (JObject)token;

            if (!obj.ContainsKey("ID") || !obj.ContainsKey("EL"))
                return null;

            var idToken = obj["ID"];
            string idString;

            if (idToken?.Type == JTokenType.Integer || idToken?.Type == JTokenType.Float)
                idString = Convert.ToString(idToken, CultureInfo.InvariantCulture); 
            else
                idString = idToken?.Value<string>() ?? "0";

            var node = new CategoryNode
            {
                Id = idString,
                Title = obj["EL"].Value<string>(),
                Parent = parent
            };

            if (obj.TryGetValue("CH", out JToken chTok) && chTok is JArray chArr)
            {
                foreach (var ch in chArr)
                {
                    var child = ParseNode(ch, node);
                    if (child != null)
                        node.Children.Add(child);
                }
            }

            return node;
        }

        private JObject ToJObject(CategoryNode n)
        {
            var obj = new JObject
            {
                ["ID"] = n.Id,  
                ["EL"] = n.Title
            };

            if (n.Children != null && n.Children.Count > 0)
            {
                var ch = new JArray();
                foreach (var c in n.Children)
                    ch.Add(ToJObject(c));
                obj["CH"] = ch;
            }

            return obj;
        }

        private IEnumerable<CategoryNode> EnumerateAll(IEnumerable<CategoryNode> start)
        {
            foreach (var n in start)
            {
                yield return n;
                foreach (var c in EnumerateAll(n.Children))
                    yield return c;
            }
        }

        private static bool TryGetLastSegment(string id, out int last)
        {
            last = 0;
            if (string.IsNullOrWhiteSpace(id)) return false;

            var parts = id.Split('.');
            if (parts.Length == 0) return false;

            string tail = parts[parts.Length - 1]; 
            return int.TryParse(tail, out last);
        }

        private string GenerateNextRootId()
        {
            int max = 0;
            foreach (var n in _root)
            {
                if (n.IsNonCatRoot) continue;
                if (int.TryParse(n.Id, out int top) && top > max) max = top;
            }
            return (max + 1).ToString();
        }

        private string GenerateNextChildId(CategoryNode parent)
        {
            int max = 0;
            foreach (var c in parent.Children)
            {
                if (TryGetLastSegment(c.Id, out int seg) && seg > max) max = seg;
            }
            return $"{parent.Id}.{max + 1}";
        }

        private string GenerateNextSiblingId(CategoryNode node)
        {
            var parent = node.Parent;
            if (parent == null)
            {
                return GenerateNextRootId();
            }
            else
            {
                int max = 0;
                foreach (var s in parent.Children)
                {
                    if (TryGetLastSegment(s.Id, out int seg) && seg > max) max = seg;
                }
                return $"{parent.Id}.{max + 1}";
            }
        }

        private CategoryNode SelectedNode => Tree.SelectedItem as CategoryNode;

        private static void ExpandCollapse(ItemsControl parent, bool expand)
        {
            foreach (var item in parent.Items)
            {
                var tvi = parent.ItemContainerGenerator.ContainerFromItem(item) as TreeViewItem;
                if (tvi != null)
                {
                    tvi.IsExpanded = expand;
                    if (tvi.Items.Count > 0)
                        ExpandCollapse(tvi, expand);
                }
            }
        }

        private void RefreshTreeBinding()
        {
            Tree.ItemsSource = null;
            Tree.ItemsSource = _root;
        }

        private void CbNames_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var name = CbNames.SelectedItem as string;
            if (name != null)
                LoadTreeForName(name);
        }

        private void BtnAddChild_Click(object sender, RoutedEventArgs e)
        {
            var sel = SelectedNode;
            if (sel == null) { MessageBox.Show("Выберите узел."); return; }
            if (sel.IsNonCatRoot) { MessageBox.Show("Данная категория защищёна и не принимает подкатегории."); return; }
            if (sel.Depth >= 3) { MessageBox.Show("Достигнут лимит вложенности"); return; }

            var child = new CategoryNode
            {
                Id = GenerateNextChildId(sel),
                Title = "Новая подкатегория",
                Parent = sel
            };
            sel.Children.Add(child);
            RefreshTreeBinding();
        }

        private void BtnAddSibling_Click(object sender, RoutedEventArgs e)
        {
            var sel = SelectedNode;

            if (sel == null)
            {
                var node = new CategoryNode { Id = GenerateNextRootId(), Title = "Новая категория" };
                _root.Add(node);
                RefreshTreeBinding();
                return;
            }

            if (sel.Parent == null)
            {
                var node = new CategoryNode { Id = GenerateNextRootId(), Title = "Новая категория" };
                _root.Add(node);
            }
            else
            {
                if (sel.Parent.IsNonCatRoot)
                {
                    MessageBox.Show("Данная категория не может иметь дочерних элементов.");
                    return;
                }

                var sib = new CategoryNode
                {
                    Id = GenerateNextSiblingId(sel),
                    Title = "Новая категория",
                    Parent = sel.Parent
                };
                sel.Parent.Children.Add(sib);
            }

            RefreshTreeBinding();
        }

        private void BtnRename_Click(object sender, RoutedEventArgs e)
        {
            var sel = SelectedNode;
            if (sel == null) { MessageBox.Show("Выберите узел."); return; }
            if (sel.IsNonCatRoot) { MessageBox.Show("Данная категория защищена от редактирования."); return; }

            var txt = PromptText("Новое имя категории:", sel.Title);
            if (!string.IsNullOrWhiteSpace(txt))
            {
                sel.Title = txt.Trim();
                RefreshTreeBinding();
            }
        }

        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            var sel = SelectedNode;
            if (sel == null) { MessageBox.Show("Выберите узел."); return; }
            if (sel.IsNonCatRoot) { MessageBox.Show("Данная категория защищена от удаления."); return; }

            string idToDelete = sel.Id;
            string titleToDelete = sel.Title;

            var confirmText = $"Удалить «{titleToDelete}» (ID: {idToDelete}) и всех его потомков? Узлы из данной категории и её наследников будут перемещены в БЕЗ КАТЕГОРИИ.";

            if (MessageBox.Show(confirmText, "Подтверждение", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                ClearNodeManagerSubcats(idToDelete);

                if (sel.Parent == null)
                    _root.Remove(sel);
                else
                    sel.Parent.Children.Remove(sel);

                RefreshTreeBinding();
            }

            SaveCurrentToDb();
        }

        private void ClearNodeManagerSubcats(string deletedId)
        {
            if (string.IsNullOrWhiteSpace(deletedId))
                return;

            try
            {
                using (var conn = OpenConn())
                using (var cmd = new SQLiteCommand(@"
                    UPDATE nodeManager
                       SET SUBCAT = 0
                     WHERE CAST(SUBCAT AS TEXT) = @id
                        OR CAST(SUBCAT AS TEXT) LIKE @prefix;", conn))
                {
                    cmd.Parameters.AddWithValue("@id", deletedId);
                    cmd.Parameters.AddWithValue("@prefix", deletedId + ".%");

                    var affected = cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка обновления субкатегорий в узлах:\n" + ex.Message, "Ошибка БД", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnExpand_Click(object sender, RoutedEventArgs e) { ExpandCollapse(Tree, true); }
        private void BtnCollapse_Click(object sender, RoutedEventArgs e) { ExpandCollapse(Tree, false); }
        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            SaveCurrentToDb();
            this.Close();
        }

        private string PromptText(string caption, string initial)
        {
            var w = new Window
            {
                Title = caption,
                Width = 400,
                Height = 160,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.NoResize,
                Owner = this,
                ShowInTaskbar = false
            };

            var grid = new Grid { Margin = new Thickness(12) };
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var tb = new TextBox { Text = initial, Margin = new Thickness(0, 0, 0, 8) };
            Grid.SetRow(tb, 0);

            var panel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var ok = new Button { Content = "OK", Width = 84, Margin = new Thickness(4) };
            var cancel = new Button { Content = "Отмена", Width = 84, Margin = new Thickness(4) };
            ok.Click += delegate { w.Tag = tb.Text; w.DialogResult = true; w.Close(); };
            cancel.Click += delegate { w.DialogResult = false; w.Close(); };

            Grid.SetRow(panel, 1);
            panel.Children.Add(ok);
            panel.Children.Add(cancel);

            grid.Children.Add(tb);
            grid.Children.Add(panel);
            w.Content = grid;

            return w.ShowDialog() == true ? (w.Tag == null ? initial : w.Tag.ToString()) : initial;
        }
    }

    public class CategoryNode
    {
        [JsonProperty("ID")]
        public string Id { get; set; }

        [JsonProperty("EL")]
        public string Title { get; set; }

        [JsonProperty("CH", NullValueHandling = NullValueHandling.Ignore)]
        public ObservableCollection<CategoryNode> Children { get; set; } = new ObservableCollection<CategoryNode>();

        [JsonIgnore]
        public CategoryNode Parent { get; set; }

        [JsonIgnore]
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

        [JsonIgnore]
        public bool IsNonCatRoot => Id == "0"; 
    }

    public sealed class BooleanToVisibilityConverter : System.Windows.Data.IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool b = value is bool && (bool)value;
            return b ? Visibility.Visible : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (value is Visibility) && ((Visibility)value == Visibility.Visible);
        }
    }
}