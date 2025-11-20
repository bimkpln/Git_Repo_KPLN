using Autodesk.Revit.DB;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media.Imaging;
using Window = System.Windows.Window;

namespace KPLN_FamilyManager.Forms
{
    public partial class FamilyManagerEditUser : Window
    {
        private const string DB_PATH = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_FamilyManager.db";
        private const string DEPT_DEF = "NOTASSIGNED";
        private const string DEPT_MAIN = "MAIN";
        private const int DEFAULT_CATEGORY_ID = 1;

        public FamilyManagerRecord ResultRecord { get; private set; }

        public string OpenFormat { get; private set; }
        public bool DeleteStatus { get; private set; }
        public string filePathFI { get; private set; }

        private string _originalStatus;

        private List<CategoryItem> _categories;
        private List<ProjectItem> _projects;
        private List<StageItem> _stages;

        public bool IsFavorite { get; set; } = false;

        public class FamilyManagerRecord
        {
            public int ID { get; set; }
            public string STATUS { get; set; }
            public string FULLPATH { get; set; }
            public string LM_DATE { get; set; }
            public int CATEGORY { get; set; }
            public int SUB_CATEGORY { get; set; }
            public int PROJECT { get; set; }
            public int STAGE { get; set; }
            public string DEPARTAMENT { get; set; }
            public string IMPORT_INFO { get; set; }
            public byte[] IMAGE { get; set; }
        }

        private class CategoryItem
        {
            public int ID { get; set; }
            public string DEPARTAMENT { get; set; }
            public string NAME { get; set; }
            public string NC_NAME { get; set; }
        }

        private class SubCategoryItem
        {
            [JsonProperty("id")]
            public int Id { get; set; }

            [JsonProperty("name")]
            public string Name { get; set; }

            [JsonProperty("dep")]
            public string DepRaw { get; set; }

            public HashSet<string> DepSet { get; set; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        private class ProjectItem
        {
            public int ID { get; set; }
            public string NAME { get; set; }
        }

        private class StageItem
        {
            public int ID { get; set; }
            public string NAME { get; set; }
        }

        public FamilyManagerEditUser(string idText)
        {
            InitializeComponent();
            EnsureFavoritesFileExists();
            LoadLookups();

            if (!int.TryParse(idText, out var id))
            {
                MessageBox.Show($"Некорректный ID: {idText}", "Family Manager", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var record = LoadRecord(id);
            if (record == null)
            {
                MessageBox.Show($"Запись с ID={id} не найдена.", "Family Manager", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            DataContext = record;

            try
            {
                var fav = IsFavoriteId(record.ID);
                IsFavorite = fav;
                if (FavoriteToggle != null) FavoriteToggle.IsChecked = fav;
            }
            catch { }

            RenderImportInfo(record.IMPORT_INFO);

            filePathFI = record.FULLPATH;

            if (!string.IsNullOrEmpty(record.FULLPATH))
            {
                Title = System.IO.Path.GetFileName(record.FULLPATH);
                FamilyName.Text = System.IO.Path.GetFileName(record.FULLPATH);
                FamilyName.ToolTip = System.IO.Path.GetFileName(record.FULLPATH);
            }

            if (record.IMAGE != null && record.IMAGE.Length > 0)
            {
                using (var ms = new MemoryStream(record.IMAGE))
                {
                    var bmp = new BitmapImage();
                    bmp.BeginInit();
                    bmp.CacheOption = BitmapCacheOption.OnLoad;
                    bmp.StreamSource = ms;
                    bmp.EndInit();
                    PreviewImage.Source = bmp;
                }
            }

            _originalStatus = record.STATUS?.Trim();
            PopulateDisplayFields(record);
        }

        // ---- ЗАГРУЗКА СПРАВОЧНИКОВ ----
        private void LoadLookups()
        {
            var cs = $"Data Source={DB_PATH};Version=3;Read Only=True;Foreign Keys=True;";

            _categories = new List<CategoryItem>();
            _projects = new List<ProjectItem>();
            _stages = new List<StageItem>();

            using (var conn = new SQLiteConnection(cs))
            {
                conn.Open();

                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"SELECT ID, DEPARTAMENT, NAME, NC_NAME FROM Category ORDER BY NAME;";
                    using (var r = cmd.ExecuteReader())
                    {
                        while (r.Read())
                        {
                            _categories.Add(new CategoryItem
                            {
                                ID = r.IsDBNull(0) ? 0 : r.GetInt32(0),
                                DEPARTAMENT = r.IsDBNull(1) ? "" : r.GetString(1),
                                NAME = r.IsDBNull(2) ? "" : r.GetString(2),
                                NC_NAME = r.IsDBNull(3) ? "" : r.GetString(3),
                            });
                        }
                    }
                }

                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"SELECT ID, NAME FROM Project ORDER BY ID;";
                    using (var r = cmd.ExecuteReader())
                    {
                        while (r.Read())
                        {
                            _projects.Add(new ProjectItem
                            {
                                ID = r.IsDBNull(0) ? 0 : r.GetInt32(0),
                                NAME = r.IsDBNull(1) ? "" : r.GetString(1),
                            });
                        }
                    }
                }

                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"SELECT ID, NAME FROM Stage ORDER BY ID;";
                    using (var r = cmd.ExecuteReader())
                    {
                        while (r.Read())
                        {
                            _stages.Add(new StageItem
                            {
                                ID = r.IsDBNull(0) ? 0 : r.GetInt32(0),
                                NAME = r.IsDBNull(1) ? "" : r.GetString(1),
                            });
                        }
                    }
                }
            }
        }

        // ---- ЗАГРУЗКА ЗАПИСИ ----
        private FamilyManagerRecord LoadRecord(int id)
        {
            var cs = $"Data Source={DB_PATH};Version=3;Read Only=True;Foreign Keys=True;";

            using (var conn = new SQLiteConnection(cs))
            {
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText =
                        @"SELECT ID, STATUS, FULLPATH, LM_DATE, CATEGORY, SUB_CATEGORY, PROJECT, STAGE,
                                 DEPARTAMENT, IMPORT_INFO, IMAGE
                          FROM FamilyManager
                          WHERE ID = @id
                          LIMIT 1;";
                    cmd.Parameters.AddWithValue("@id", id);

                    using (var reader = cmd.ExecuteReader(CommandBehavior.SingleRow))
                    {
                        if (!reader.Read()) return null;

                        return new FamilyManagerRecord
                        {
                            ID = reader.GetInt32(0),
                            STATUS = reader.IsDBNull(1) ? "" : reader.GetString(1),
                            FULLPATH = reader.IsDBNull(2) ? "" : reader.GetString(2),
                            LM_DATE = reader.IsDBNull(3) ? "" : reader.GetString(3),
                            CATEGORY = reader.IsDBNull(4) ? 0 : reader.GetInt32(4),
                            SUB_CATEGORY = reader.IsDBNull(5) ? 0 : reader.GetInt32(5),
                            PROJECT = reader.IsDBNull(6) ? 0 : reader.GetInt32(6),
                            STAGE = reader.IsDBNull(7) ? 0 : reader.GetInt32(7),
                            DEPARTAMENT = reader.IsDBNull(8) ? "" : reader.GetString(8),
                            IMPORT_INFO = reader.IsDBNull(9) ? "" : reader.GetString(9),
                            IMAGE = reader.IsDBNull(10) ? null : (byte[])reader[10]
                        };
                    }
                }
            }
        }

        // ---- РЕНДЕР ИМПОРТ-ИНФО ----
        private void RenderImportInfo(string json)
        {
            ImportInfoBox.Document.Blocks.Clear();
            if (string.IsNullOrWhiteSpace(json)) return;

            try
            {
                var obj = JObject.Parse(json);
                var para = new Paragraph { Margin = new Thickness(0) };
                foreach (var prop in obj.Properties())
                {
                    para.Inlines.Add(new Bold(new Run($"{prop.Name}: ")));
                    var val = prop.Value?.Type == JTokenType.Null ? "—" : (prop.Value?.ToString() ?? "—");
                    para.Inlines.Add(new Run(val));
                    para.Inlines.Add(new LineBreak());
                }
                ImportInfoBox.Document.Blocks.Add(para);
            }
            catch
            {
                ImportInfoBox.Document.Blocks.Add(new Paragraph(new Run(json)) { Margin = new Thickness(0) });
            }
        }

        // ---- ПРЕОБРАЗОВАНИЕ SUBCATEGORY JSON ----
        private static List<SubCategoryItem> ParseSubcategories(string json)
        {
            var fallback = new List<SubCategoryItem> { new SubCategoryItem { Id = 0, Name = "Корневая дирректория" } };
            if (string.IsNullOrWhiteSpace(json)) return fallback;

            try
            {
                var items = JsonConvert.DeserializeObject<List<SubCategoryItem>>(json);
                if (items == null || items.Count == 0) return fallback;

                foreach (var it in items)
                {
                    it.DepSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    var raw = it.DepRaw ?? "";
                    foreach (var part in raw.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        var p = part.Trim();
                        if (string.IsNullOrEmpty(p)) continue;
                        if (p.Equals(DEPT_DEF, StringComparison.OrdinalIgnoreCase)) continue;
                        if (p.Equals(DEPT_MAIN, StringComparison.OrdinalIgnoreCase)) continue;
                        it.DepSet.Add(p);
                    }
                }

                return items
                    .Where(i => i != null)
                    .GroupBy(i => i.Id)
                    .Select(g => g.First())
                    .OrderBy(i => i.Id)
                    .ToList();
            }
            catch
            {
                return fallback;
            }
        }

        // ---- ТЕКСТОВЫЕ ПОЛЯ ----
        private void PopulateDisplayFields(FamilyManagerRecord rec)
        {
            DeptValue.Text = string.IsNullOrWhiteSpace(rec.DEPARTAMENT) ? "—" : rec.DEPARTAMENT;

            var cat = _categories.FirstOrDefault(c => c.ID == rec.CATEGORY)
                   ?? _categories.FirstOrDefault(c => c.ID == DEFAULT_CATEGORY_ID)
                   ?? _categories.OrderBy(c => c.ID).FirstOrDefault();

            CategoryValue.Text = cat?.NAME ?? "Не назначен";

            // Подкатегория (если > 0 и есть имя)
            var subName = "";
            if (cat != null)
            {
                var subcats = ParseSubcategories(cat.NC_NAME);
                var sub = subcats.FirstOrDefault(s => s.Id == rec.SUB_CATEGORY);
                if (sub != null && sub.Id != 0 && !string.IsNullOrWhiteSpace(sub.Name))
                {
                    subName = sub.Name;
                }
            }

            if (!string.IsNullOrWhiteSpace(subName))
            {
                CategorySep.Visibility = System.Windows.Visibility.Visible;
                SubCategoryValue.Visibility = System.Windows.Visibility.Visible;
                SubCategoryValue.Text = subName;
            }
            else
            {
                CategorySep.Visibility = System.Windows.Visibility.Collapsed;
                SubCategoryValue.Visibility = System.Windows.Visibility.Collapsed;
                SubCategoryValue.Text = "";
            }

            var prj = _projects.FirstOrDefault(p => p.ID == rec.PROJECT);
            ProjectValue.Text = prj?.NAME ?? "—";
            var stg = _stages.FirstOrDefault(s => s.ID == rec.STAGE);
            StageValue.Text = stg?.NAME ?? "—";
        }

        // ---- ОТКРЫТЬ ПАПКУ ----
        private void BtnOpenFolder_Click(object sender, RoutedEventArgs e)
        {
            var rec = DataContext as FamilyManagerRecord;
            if (rec == null || string.IsNullOrEmpty(rec.FULLPATH))
            {
                MessageBox.Show("Путь к файлу не задан.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                string filePath = rec.FULLPATH;

                if (File.Exists(filePath))
                {
                    Process.Start("explorer.exe", "/select,\"" + filePath + "\"");
                    return;
                }

                string dir = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir))
                {
                    Process.Start("explorer.exe", "\"" + dir + "\"");
                }
                else
                {
                    MessageBox.Show("Файл или папка не найдены.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при открытии папки:\n" + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnCancel_OnClick(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void ButtonOpenFamily_Click(object sender, RoutedEventArgs e)
        {
            OpenFormat = "OpenFamily";
            DialogResult = true;
            Close();
        }

        private void ButtonOpenFamilyInProject_Click(object sender, RoutedEventArgs e)
        {
            OpenFormat = "OpenFamilyInProject";
            DialogResult = true;
            Close();
        }

        // ---- ИЗБРАННОЕ ----
        private void FavoriteToggle_Click(object sender, RoutedEventArgs e)
        {
            var rec = DataContext as FamilyManagerRecord;
            if (rec == null) return;

            var on = FavoriteToggle.IsChecked == true;
            IsFavorite = on;
            SetFavorite(rec.ID, on);
        }

        private static string GetFavoritesFilePath()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "RevitFamilyManagerFavorites.txt");
        }

        private static void EnsureFavoritesFileExists()
        {
            var path = GetFavoritesFilePath();
            var dir = Path.GetDirectoryName(path);
            try
            {
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                if (!File.Exists(path))
                {
                    using (var fs = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read))
                    using (var sw = new StreamWriter(fs, Encoding.UTF8)) { }
                }
            }
            catch { }
        }

        private static HashSet<int> LoadFavorites()
        {
            EnsureFavoritesFileExists();
            var set = new HashSet<int>();
            var path = GetFavoritesFilePath();
            try
            {
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var sr = new StreamReader(fs, Encoding.UTF8, true))
                {
                    var text = sr.ReadToEnd();
                    foreach (var t in (text ?? "").Split(new[] { ',', ';', ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
                        if (int.TryParse(t.Trim(), out var id)) set.Add(id);
                }
            }
            catch { }
            return set;
        }

        private static void SaveFavorites(HashSet<int> set)
        {
            EnsureFavoritesFileExists();

            var path = GetFavoritesFilePath();
            var tmp = path + ".tmp";
            var payload = string.Join(",", set.OrderBy(v => v));

            try
            {
                File.WriteAllText(tmp, payload, Encoding.UTF8);
                if (File.Exists(path)) File.Replace(tmp, path, null);
                else File.Move(tmp, path);
            }
            catch
            {
                try { if (File.Exists(tmp)) File.Delete(tmp); } catch { }
            }
        }

        private static void SetFavorite(int id, bool isFav)
        {
            var set = LoadFavorites();
            if (isFav) set.Add(id); else set.Remove(id);
            SaveFavorites(set);
        }

        private static bool IsFavoriteId(int id)
        {
            var set = LoadFavorites();
            return set.Contains(id);
        }
    }
}
