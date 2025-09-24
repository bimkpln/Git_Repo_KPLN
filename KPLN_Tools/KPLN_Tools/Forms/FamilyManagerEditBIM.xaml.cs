using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Windows.Documents;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

using CheckBox = System.Windows.Controls.CheckBox;
using Style = System.Windows.Style;
using Window = System.Windows.Window;
using System.Windows.Controls;


namespace KPLN_Tools.Forms
{
    public partial class FamilyManagerEditBIM : Window
    {
        private const string DB_PATH = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_FamilyManager.db";
        private static readonly string[] ALL_DEPARTMENTS = { "АР", "КР", "ОВиК", "ВК", "ЭОМ", "СС", "BIM" };

        private const string DEPT_DEF = "NOTASSIGNED";
        private const string DEPT_MAIN = "MAIN";
        private const int DEFAULT_CATEGORY_ID = 1;

        private const int TARGET_MAX_SIDE = 512;   
        private const long JPEG_QUALITY = 85L; 
        private byte[] _preparedImageBytes;

        private bool _isEditingImportInfo = false;
        private string _importInfoBackup = null;

        public FamilyManagerRecord ResultRecord { get; private set; }

        public string OpenFormat { get; private set; }
        public bool DeleteStatus { get; private set; }
        public string filePathFI { get; private set; }

        private string _originalStatus;

        private List<CategoryItem> _categories;
        private List<ProjectItem> _projects;
        private List<StageItem> _stages;

        // Класс БД. Общее
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

        // Класс БД. Категории
        private class CategoryItem
        {
            public int ID { get; set; }
            public string DEPARTAMENT { get; set; }
            public string NAME { get; set; }
            public string NC_NAME { get; set; }
        }

        // Класс JSON. Подкатегории
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

        // Класс БД. Проекты
        private class ProjectItem
        {
            public int ID { get; set; }
            public string NAME { get; set; }
        }

        // Класс БД. Стадии
        private class StageItem
        {
            public int ID { get; set; }
            public string NAME { get; set; }
        }

        // Основной конструктор
        public FamilyManagerEditBIM(string idText)
        {
            InitializeComponent();

            LoadLookups(); // Справочники

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
            BuildDepartmentUI(record);
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
            SetupStatusCombo(_originalStatus); // Статусы

            ProjectCombo.ItemsSource = _projects;
            ProjectCombo.DisplayMemberPath = "NAME";
            ProjectCombo.SelectedValuePath = "ID";
            ProjectCombo.SelectedValue = record.PROJECT;

            StageCombo.ItemsSource = _stages;
            StageCombo.DisplayMemberPath = "NAME";
            StageCombo.SelectedValuePath = "ID";
            StageCombo.SelectedValue = record.STAGE;

            RefreshCategoryCombos(); // record.DEPARTAMENT внутри
            CategoryCombo.SelectionChanged += CategoryCombo_SelectionChanged;
        }

        // БД. Category, Project, Stage
        private void LoadLookups()
        {
            var cs = $"Data Source={DB_PATH};Version=3;Read Only=True;Foreign Keys=True;";

            _categories = new List<CategoryItem>();
            _projects = new List<ProjectItem>();
            _stages = new List<StageItem>();

            using (var conn = new SQLiteConnection(cs))
            {
                conn.Open();

                // Category (ID, NAME, NC_NAME)
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

                // Project (ID, NAME)
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

                // Stage (ID, NAME)
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

        // БД. FamilyManager
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

        // Создание ComboBox со статусом
        private void SetupStatusCombo(string status)
        {
            StatusCombo.Items.Clear();
            StatusCombo.IsEnabled = true;

            var s = (status ?? "").Trim().ToUpperInvariant();

            switch (s)
            {
                case "NEW":
                    StatusCombo.Items.Add("Подтверждено");
                    StatusCombo.Items.Add("Требует подтверждения");
                    StatusCombo.Items.Add("Игнорировать в БД");
                    StatusCombo.SelectedItem = "Требует подтверждения";
                    ButtonSaveDB.IsEnabled = true;
                    break;

                case "OK":
                    StatusCombo.Items.Add("Подтверждено");
                    StatusCombo.Items.Add("Игнорировать в БД");
                    StatusCombo.SelectedItem = "Подтверждено";
                    ButtonSaveDB.IsEnabled = true;
                    break;

                case "IGNORE":
                    StatusCombo.Items.Add("Подтверждено");
                    StatusCombo.Items.Add("Игнорировать в БД");
                    StatusCombo.SelectedItem = "Игнорировать в БД";
                    ButtonSaveDB.IsEnabled = true;
                    break;

                case "ABSENT":
                    StatusCombo.Items.Add("Файл не найден на диске");
                    StatusCombo.SelectedItem = "Файл не найден на диске";
                    StatusCombo.IsEnabled = false;
                    ButtonDeleteDB.IsEnabled = true;
                    break;

                case "ERROR":
                    StatusCombo.Items.Add("В записи БД присутствует ошибка");
                    StatusCombo.SelectedItem = "В записи БД присутствует ошибка";
                    StatusCombo.IsEnabled = false;
                    ButtonDeleteDB.IsEnabled = true;
                    break;

                default:
                    StatusCombo.Items.Add("Не удалось считать статус");
                    StatusCombo.SelectedItem = "Не удалось считать статус";
                    StatusCombo.IsEnabled = false;
                    break;
            }
        }


        // Парсинг строки "Отдел"
        private static HashSet<string> ParseDepartments(string s)
        {
            return (s ?? string.Empty)
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Trim())
                .Where(x => !string.IsNullOrEmpty(x))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
        }

        // Сборка строки "Отдел" обратно в фиксированном порядке
        private static string JoinDepartments(IEnumerable<string> selected)
        {
            var set = new HashSet<string>(selected ?? Array.Empty<string>(), StringComparer.Ordinal);
            var ordered = ALL_DEPARTMENTS.Where(set.Contains);
            return string.Join(", ", ordered);
        }

        // Построение UI чекбоксов "Отдел"
        private void BuildDepartmentUI(FamilyManagerRecord rec)
        {
            DeptPanel.Children.Clear();

            var selected = ParseDepartments(rec?.DEPARTAMENT);
            foreach (var dep in ALL_DEPARTMENTS)
            {
                var cb = new System.Windows.Controls.CheckBox
                {
                    Content = dep,
                    IsChecked = selected.Contains(dep),
                };
                cb.Style = (Style)FindResource("DeptCheckStyle");
                cb.Checked += DepartmentCheckChanged;
                cb.Unchecked += DepartmentCheckChanged;

                DeptPanel.Children.Add(cb);
            }

            UpdateDeptToggleVisual();
        }
      
        // Категории, прошедшие фильтр по отделам
        private static List<CategoryItem> FilterCategoriesByDepartments(
            IEnumerable<CategoryItem> all, string selectedDepartments)
        {
            var selRaw = ParseDepartments(selectedDepartments);
            var sel = NormalizeDeptSet(selRaw);
            int selCnt = sel.Count;

            return all.Where(c =>
            {
                var cdepsRaw = ParseDepartments(c.DEPARTAMENT);

                if (cdepsRaw.Any(d => d.Equals(DEPT_MAIN, StringComparison.OrdinalIgnoreCase))) return true;
                if (cdepsRaw.Any(d => d.Equals(DEPT_DEF, StringComparison.OrdinalIgnoreCase))) return true;

                var cdeps = NormalizeDeptSet(cdepsRaw);

                if (selCnt == 0)
                    return cdepsRaw.Count == 0;

                return sel.All(s => cdeps.Contains(s));
            })
            .OrderBy(c => IsDefaultCategory(c) ? 0 : 1)                             
            .ThenBy(c => c.NAME, StringComparer.CurrentCultureIgnoreCase)
            .ToList();
        }

        // Вспомогательный метод сортировки
        private static bool IsDefaultCategory(CategoryItem c)
        {
            if (c == null) return false;
            return c.ID == DEFAULT_CATEGORY_ID
                || string.Equals(c.NAME, "Не назначен", StringComparison.CurrentCultureIgnoreCase);
        }

        // Убираем служебные и пустые отделы перед сравнением
        private static HashSet<string> NormalizeDeptSet(IEnumerable<string> deps)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var d in deps ?? Enumerable.Empty<string>())
            {
                if (string.IsNullOrWhiteSpace(d)) continue;
                if (string.Equals(d, DEPT_DEF, StringComparison.OrdinalIgnoreCase)) continue;
                if (string.Equals(d, DEPT_MAIN, StringComparison.OrdinalIgnoreCase)) continue;
                set.Add(d.Trim());
            }
            return set;
        }

        // Перерасчёт категорий из отделов
        private void RefreshCategoryCombos()
        {
            var record = DataContext as FamilyManagerRecord;
            if (record == null) return;

            var filtered = FilterCategoriesByDepartments(_categories, record.DEPARTAMENT);

            if (filtered.Count == 0)
            {
                var def = _categories.FirstOrDefault(c => c.ID == 1) ?? _categories.OrderBy(c => c.ID).FirstOrDefault();
                if (def != null)
                {
                    filtered = new List<CategoryItem> { def };
                    record.CATEGORY = def.ID;
                }
            }

            int? prevId = CategoryCombo.SelectedValue as int?;
            bool keepPrev = prevId.HasValue && filtered.Any(c => c.ID == prevId.Value);

            CategoryCombo.ItemsSource = filtered;
            CategoryCombo.DisplayMemberPath = "NAME";
            CategoryCombo.SelectedValuePath = "ID";
            CategoryCombo.SelectedValue = keepPrev ? prevId.Value : (object)record.CATEGORY;

            if (!(CategoryCombo.SelectedValue is int))
            {
                var def = _categories.FirstOrDefault(c => c.ID == 1) ?? _categories.OrderBy(c => c.ID).FirstOrDefault();
                if (def != null)
                {
                    record.CATEGORY = def.ID;
                    CategoryCombo.SelectedValue = def.ID;
                }
            }
            else
            {
                record.CATEGORY = (int)CategoryCombo.SelectedValue;
            }

            var currentCat = filtered.FirstOrDefault(c => c.ID == record.CATEGORY)
                          ?? filtered.FirstOrDefault()
                          ?? (_categories.FirstOrDefault(c => c.ID == 1) ?? _categories.OrderBy(c => c.ID).FirstOrDefault());

            var subcatsAll = ParseSubcategories(currentCat?.NC_NAME);
            var subcats = FilterSubcategoriesByDept(subcatsAll, record.DEPARTAMENT);

            if (!subcats.Any(sc => sc.Id == record.SUB_CATEGORY))
            {
                record.SUB_CATEGORY = subcats.FirstOrDefault()?.Id ?? 0;
            }

            CategoryNcCombo.ItemsSource = subcats;
            CategoryNcCombo.DisplayMemberPath = "Name";
            CategoryNcCombo.SelectedValuePath = "Id";
            CategoryNcCombo.SelectedValue = record.SUB_CATEGORY;
        }

        // Парсинг Category.NC_NAME (JSON) → List<SubCategoryItem>
        private static List<SubCategoryItem> ParseSubcategories(string json)
        {
            var fallback = new List<SubCategoryItem>
            {
                new SubCategoryItem { Id = 0, Name = "Корневая дирректория" }
            };

            if (string.IsNullOrWhiteSpace(json))
                return fallback;

            try
            {
                var items = Newtonsoft.Json.JsonConvert
                    .DeserializeObject<List<SubCategoryItem>>(json);

                if (items == null || items.Count == 0)
                    return fallback;

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

        // Фильтр субкатегорий
        private static List<SubCategoryItem> FilterSubcategoriesByDept(List<SubCategoryItem> subcats, string departmentsCsv)
        {
            var selected = NormalizeDeptSet(ParseDepartments(departmentsCsv));
            int cnt = selected.Count;

            IEnumerable<SubCategoryItem> filtered;

            if (cnt == 0)
            {
                filtered = subcats.Where(sc => sc.DepSet == null || sc.DepSet.Count == 0);
            }
            else if (cnt == 1)
            {
                string one = selected.First();
                filtered = subcats.Where(sc =>
                    sc.DepSet == null || sc.DepSet.Count == 0 || sc.DepSet.Contains(one));
            }
            else
            {
                filtered = subcats.Where(sc =>
                    sc.DepSet == null || sc.DepSet.Count == 0 || selected.All(d => sc.DepSet.Contains(d)));
            }

            return filtered
                .OrderBy(sc => sc.Id == 0 ? 0 : 1) // сначала "Корневая директория"
                .ThenBy(sc => sc.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        // Получение байтов из Image.Source
        private static byte[] GetImageBytesFromImageControl(System.Windows.Controls.Image img)
        {
            if (img?.Source is BitmapSource bmpSrc)
            {
                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(bmpSrc));
                using (var ms = new MemoryStream())
                {
                    encoder.Save(ms);
                    return ms.ToArray();
                }
            }
            return null;
        }

        // Чтение Bitmap
        private static Bitmap LoadBitmapUnlocked(string path)
        {
            var ext = System.IO.Path.GetExtension(path)?.ToLowerInvariant();
            var ok = ext == ".png" || ext == ".bmp" || ext == ".gif";
            if (!ok)
                throw new NotSupportedException($"Формат файла не поддерживается GDI+: {ext ?? "без расширения"}");

            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var ms = new MemoryStream())
            {
                fs.CopyTo(ms);
                ms.Position = 0;
                return new Bitmap(ms);
            }
        }

        // Редактирование ориентации
        private static void FixOrientationIfNeeded(Bitmap bmp)
        {
            const int ExifOrientationId = 0x0112;
            if (!bmp.PropertyIdList.Contains(ExifOrientationId)) return;

            try
            {
                var prop = bmp.GetPropertyItem(ExifOrientationId);
                int val = BitConverter.ToUInt16(prop.Value, 0);
                RotateFlipType flip = RotateFlipType.RotateNoneFlipNone;

                switch (val)
                {
                    case 2: flip = RotateFlipType.RotateNoneFlipX; break;
                    case 3: flip = RotateFlipType.Rotate180FlipNone; break;
                    case 4: flip = RotateFlipType.Rotate180FlipX; break;
                    case 5: flip = RotateFlipType.Rotate90FlipX; break;
                    case 6: flip = RotateFlipType.Rotate90FlipNone; break;
                    case 7: flip = RotateFlipType.Rotate270FlipX; break;
                    case 8: flip = RotateFlipType.Rotate270FlipNone; break;
                    default: return;
                }

                bmp.RotateFlip(flip);

                bmp.RemovePropertyItem(ExifOrientationId);
            }
            catch { }
        }

        // Bitmap -> JPEG bytes
        private static byte[] ToJpegBytesSafe(Bitmap bmp, long quality)
        {
            using (var compatible = MakeJpegCompatible(bmp))
            using (var ms = new MemoryStream())
            {
                try
                {
                    var codec = ImageCodecInfo.GetImageDecoders()
                        .FirstOrDefault(c => c.FormatID == ImageFormat.Jpeg.Guid);

                    if (codec != null)
                    {
                        using (var encParams = new EncoderParameters(1))
                        {
                            encParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, quality);
                            compatible.Save(ms, codec, encParams);
                        }
                    }
                    else
                    {
                        compatible.Save(ms, ImageFormat.Jpeg);
                    }
                }
                catch
                {
                    ms.SetLength(0);
                    compatible.Save(ms, ImageFormat.Jpeg);
                }
                return ms.ToArray();
            }
        }

        // Приведение к формату, совместимому с JPEG
        private static Bitmap MakeJpegCompatible(Bitmap src)
        {
            var dest = new Bitmap(src.Width, src.Height, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
            dest.SetResolution(src.HorizontalResolution, src.VerticalResolution);
            using (var g = Graphics.FromImage(dest))
            {
                g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                g.DrawImage(src, new System.Drawing.Rectangle(0, 0, src.Width, src.Height));
            }
            return dest;
        }

        // Показ превью в WPF из байтов 
        private void SetPreviewFromBytes(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0)
            {
                PreviewImage.Source = null;
                return;
            }
            var bi = new BitmapImage();
            using (var ms = new MemoryStream(bytes))
            {
                bi.BeginInit();
                bi.CacheOption = BitmapCacheOption.OnLoad;
                bi.StreamSource = ms;
                bi.EndInit();
            }
            PreviewImage.Source = bi;
        }

        // Универсальная подготовка пользовательской картинки
        private byte[] PrepareUserImageFromFile(string filePath)
        {
            try
            {
                using (var src = LoadBitmapUnlocked(filePath))
                {
                    FixOrientationIfNeeded(src);
                    using (var resized = ResizeKeepingAspect(src, TARGET_MAX_SIDE))
                    {
                        return ToJpegBytesSafe(resized, JPEG_QUALITY);
                    }
                }
            }
            catch (NotSupportedException ns)
            {
                MessageBox.Show(ns.Message, "Family Manager", MessageBoxButton.OK, MessageBoxImage.Warning);
                return null;
            }
            catch (ArgumentException)
            {
                MessageBox.Show("Файл не распознан как изображение (возможно, повреждён).", "Family Manager",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return null;
            }
        }

        // Универсальная подготовка превью из .rfa
        private byte[] PrepareRfaPreviewBytes(string rfaPath)
        {
            using (var sh = GetShellThumbnail(rfaPath, TARGET_MAX_SIDE))
            {
                if (sh == null) return null;
                using (var resized = ResizeKeepingAspect(sh, TARGET_MAX_SIDE))
                {
                    return ToJpegBytesSafe(resized, JPEG_QUALITY);
                }
            }
        }

        // Ресайз с сохранением пропорций
        private static Bitmap ResizeKeepingAspect(Bitmap src, int maxSide)
        {
            double scale = Math.Min((double)maxSide / src.Width, (double)maxSide / src.Height);
            int w = Math.Max(1, (int)Math.Round(src.Width * scale));
            int h = Math.Max(1, (int)Math.Round(src.Height * scale));

            var dest = new Bitmap(w, h, PixelFormat.Format32bppArgb);
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

        // Shell thumbnail для .rfa
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void SHCreateItemFromParsingName(
            [MarshalAs(UnmanagedType.LPWStr)] string pszPath,
            IntPtr pbc,
            [MarshalAs(UnmanagedType.LPStruct)] Guid riid,
            out IShellItemImageFactory ppv);

        private static readonly Guid IID_IShellItemImageFactory = new Guid("bcc18b79-ba16-442f-80c4-8a59c30c463b");

        [ComImport]
        [Guid("bcc18b79-ba16-442f-80c4-8a59c30c463b")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellItemImageFactory
        {
            void GetImage(SIZE size, SIIGBF flags, out IntPtr phbm);
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct SIZE { public int cx; public int cy; }

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

                using (var temp = System.Drawing.Image.FromHbitmap(hBmp))
                {
                    return new Bitmap(temp);
                }
            }
            finally
            {
                if (hBmp != IntPtr.Zero) DeleteObject(hBmp);
                if (factory != null) Marshal.ReleaseComObject(factory);
            }
        }

        // Мапинг статуса
        private string MapStatusForSave(string originalStatus, string selectedUiStatus)
        {
            if (selectedUiStatus == "Требует подтверждения" ||
                selectedUiStatus == "Файл не найден на диске" ||
                selectedUiStatus == "В записи БД присутствует ошибка")
            {
                return originalStatus;
            }

            if (selectedUiStatus == "Игнорировать в БД") return "IGNORE";
            if (selectedUiStatus == "Подтверждено") return "OK";

            return originalStatus;
        }

        // IMPORT_INFO. Парсинг JSON
        private void RenderImportInfo(string json)
        {
            ImportInfoBox.Document.Blocks.Clear();

            if (string.IsNullOrWhiteSpace(json))
                return;

            try
            {
                var obj = JObject.Parse(json);

                var para = new Paragraph { Margin = new Thickness(0) };

                foreach (var prop in obj.Properties())
                {
                    para.Inlines.Add(new Bold(new Run($"{prop.Name}: ")));

                    var val = prop.Value?.Type == JTokenType.Null
                        ? "—"
                        : (prop.Value?.ToString() ?? "—");

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

        // IMPORT_INFO. Кол-во отделов выбрано
        private int GetSelectedDepartmentsCount()
        {
            return DeptPanel.Children
                .OfType<CheckBox>()
                .Count(c => c.IsChecked == true);
        }

        // IMPORT_INFO. Вытащить сырой текст из RichTextBox
        private static string GetRtbText(RichTextBox rtb)
        {
            var tr = new TextRange(rtb.Document.ContentStart, rtb.Document.ContentEnd);
            return tr.Text?.TrimEnd('\r', '\n');
        }

        // IMPORT_INFO. Показать сырой JSON для редактирования (моноширинный)
        private void ShowRawJsonInEditor(string json)
        {
            ImportInfoBox.Document.Blocks.Clear();
            var para = new Paragraph { Margin = new Thickness(0) };
            para.Inlines.Add(new Run(json ?? string.Empty)
            {
                FontFamily = new System.Windows.Media.FontFamily("Consolas")
            });
            ImportInfoBox.Document.Blocks.Add(para);
        }

        // IMPORT_INFO. Проверка и нормализация JSON 
        private static string ValidateAndNormalizeJson(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                throw new Exception("JSON пустой.");

            var token = JToken.Parse(raw);
            if (token.Type != JTokenType.Object)
                throw new Exception("Ожидался JSON-объект { ... }.");

            var obj = (JObject)token;

            foreach (var p in obj.Properties().ToList())
            {
                if (p.Value.Type == JTokenType.Null) p.Value = "";
            }

            return obj.ToString(Formatting.None);
        }

        // Статистический метод записи данных в БД
        public static void SaveRecordToDatabase(FamilyManagerRecord rec)
        {
            if (rec == null) throw new ArgumentNullException(nameof(rec));

            // NB: тут важно Read Only = False
            var cs = $"Data Source={DB_PATH};Version=3;Read Only=False;Foreign Keys=True;";

            using (var conn = new SQLiteConnection(cs))
            {
                conn.Open();
                using (var tr = conn.BeginTransaction())
                using (var cmd = conn.CreateCommand())
                {
                    cmd.Transaction = tr;

                    cmd.CommandText = @"
                        UPDATE FamilyManager
                        SET
                            STATUS = @status,
                            FULLPATH = @fullpath,
                            LM_DATE = @lmdate,
                            CATEGORY = @category,
                            SUB_CATEGORY = @subcat,  
                            PROJECT = @project,
                            STAGE = @stage,
                            DEPARTAMENT = @dept,
                            IMPORT_INFO = @import,
                            IMAGE = @image
                        WHERE ID = @id;";

                    cmd.Parameters.AddWithValue("@id", rec.ID);
                    cmd.Parameters.AddWithValue("@status", (object)(rec.STATUS ?? string.Empty));
                    cmd.Parameters.AddWithValue("@fullpath", (object)(rec.FULLPATH ?? string.Empty));
                    cmd.Parameters.AddWithValue("@lmdate", (object)(rec.LM_DATE ?? string.Empty));
                    cmd.Parameters.AddWithValue("@category", rec.CATEGORY);
                    cmd.Parameters.AddWithValue("@subcat", rec.SUB_CATEGORY);
                    cmd.Parameters.AddWithValue("@project", rec.PROJECT);
                    cmd.Parameters.AddWithValue("@stage", rec.STAGE);
                    cmd.Parameters.AddWithValue("@dept", (object)(rec.DEPARTAMENT ?? string.Empty));
                    cmd.Parameters.AddWithValue("@import", (object)(rec.IMPORT_INFO ?? string.Empty));

                    var pImage = cmd.CreateParameter();
                    pImage.ParameterName = "@image";
                    pImage.DbType = System.Data.DbType.Binary;
                    pImage.Value = (object)rec.IMAGE ?? DBNull.Value;
                    cmd.Parameters.Add(pImage);

                    var affected = cmd.ExecuteNonQuery();
                    tr.Commit();
                }
            }
        }

        // XAML. Открытие папки с семейством
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

        // XAML. Чек-боксы "Отдел"
        private void DepartmentCheckChanged(object sender, RoutedEventArgs e)
        {
            var record = DataContext as FamilyManagerRecord;
            if (record == null) return;

            var selected = DeptPanel.Children
                .OfType<System.Windows.Controls.CheckBox>()
                .Where(c => c.IsChecked == true)
                .Select(c => c.Content?.ToString())
                .Where(s => !string.IsNullOrWhiteSpace(s));

            record.DEPARTAMENT = JoinDepartments(selected);
            RefreshCategoryCombos();
            UpdateDeptToggleVisual();
        }

        // XAML. Кнопка выбрать/снять все отделы
        private void DeptToggleBtn_Click(object sender, RoutedEventArgs e)
        {
            bool allChecked = DeptPanel.Children.OfType<CheckBox>().All(cb => cb.IsChecked == true);
            SetAllDepartments(!allChecked);
            UpdateDeptToggleVisual();
        }

        // Выбрать все чекбоксы с оделами
        private void SetAllDepartments(bool check)
        {
            foreach (var cb in DeptPanel.Children.OfType<CheckBox>())
                cb.IsChecked = check;

            UpdateDeptToggleVisual();
        }

        // Обновления статуса кнопки для выбора всех отделов
        private void UpdateDeptToggleVisual()
        {
            if (DeptToggleBtn == null) return;

            bool allChecked = DeptPanel.Children.OfType<CheckBox>().All(cb => cb.IsChecked == true);
            DeptToggleBtn.Content = allChecked ? "Снять все" : "Выбрать все";
            DeptToggleBtn.ToolTip = allChecked ? "Снять все отделы" : "Выбрать все отделы";
        }

        // XAML.Чек-боксы отделов
        private void CategoryCombo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var record = DataContext as FamilyManagerRecord;
            if (record == null) return;

            if (CategoryCombo.SelectedValue is int catId)
                record.CATEGORY = catId;

            var filtered = CategoryCombo.ItemsSource as IEnumerable<CategoryItem> ?? Enumerable.Empty<CategoryItem>();
            var currentCat = filtered.FirstOrDefault(c => c.ID == record.CATEGORY);
            var subcatsAll = ParseSubcategories(currentCat?.NC_NAME);
            var subcats = FilterSubcategoriesByDept(subcatsAll, record.DEPARTAMENT);

            if (!subcats.Any(sc => sc.Id == record.SUB_CATEGORY))
                record.SUB_CATEGORY = subcats.FirstOrDefault()?.Id ?? 0;

            CategoryNcCombo.ItemsSource = subcats;
            CategoryNcCombo.SelectedValue = record.SUB_CATEGORY;
        }

        // XAML. Пользовательское изображение
        private void ButtonUserPic_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new Microsoft.Win32.OpenFileDialog
                {
                    Title = "Выберите изображение",
                    Filter = "Картинки (png, gif, bmp)|*.png;*.bmp;*.gif",
                    CheckFileExists = true,
                    Multiselect = false
                };
                if (dlg.ShowDialog(this) != true)
                    return;

                // готовим JPEG нужного размера
                var jpegBytes = PrepareUserImageFromFile(dlg.FileName);
                if (jpegBytes == null || jpegBytes.Length == 0)
                {
                    MessageBox.Show("Не удалось обработать изображение.", "Family Manager",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                SetPreviewFromBytes(jpegBytes);
                _preparedImageBytes = jpegBytes;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при обработке изображения: " + ex.Message, "Family Manager",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // XAML. Изображение из RFA
        private void ButtonRFAPic_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var record = DataContext as FamilyManagerRecord;
                string path = record?.FULLPATH;

                if (!IsValidRfaOrRvtPath(path))
                {
                    var dlg = new Microsoft.Win32.OpenFileDialog
                    {
                        Title = "Выберите семейство Revit (*.rfa)",
                        Filter = "Revit Family (*.rfa)|*.rfa|Все файлы|*.*",
                        CheckFileExists = true,
                        Multiselect = false
                    };
                    if (dlg.ShowDialog(this) == true)
                        path = dlg.FileName;
                }

                if (!IsValidRfaOrRvtPath(path))
                {
                    MessageBox.Show("Путь к файлу RFA не задан или файл не найден.", "Family Manager",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var jpegBytes = PrepareRfaPreviewBytes(path);
                if (jpegBytes == null || jpegBytes.Length == 0)
                {
                    MessageBox.Show("Не удалось получить превью из файла RFA.", "Family Manager",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                SetPreviewFromBytes(jpegBytes);
                _preparedImageBytes = jpegBytes;

                if (record != null && string.IsNullOrWhiteSpace(record.FULLPATH))
                    record.FULLPATH = path;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при получении превью: " + ex.Message, "Family Manager",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
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

        // XAML. Отмена
        private void BtnCancel_OnClick(object sender, RoutedEventArgs e)
        {
            Close();
        }

        // XAML. Открыть семейство
        private void ButtonOpenFamily_Click(object sender, RoutedEventArgs e)
        {
            OpenFormat = "OpenFamily";
            DialogResult = true;
            Close();
        }

        // XAML. Загрузить семейство в проект
        private void ButtonOpenFamilyInProject_Click(object sender, RoutedEventArgs e)
        {
            OpenFormat = "OpenFamilyInProject";
            DialogResult = true;
            Close();
        }

        // XAML. Обновить информацию о семействе
        private void ButtonUpdateFamilyInfo_Click(object sender, RoutedEventArgs e)
        {
            OpenFormat = "UpdateFamily";
            DialogResult = true;
            Close();
        }

        // XAML. Редактировать информацию о семействе
        private void ButtonEditFamilyInfo_Click(object sender, RoutedEventArgs e)
        {
            var record = DataContext as FamilyManagerRecord;
            if (record == null) return;

            if (!_isEditingImportInfo)
            {
                if (GetSelectedDepartmentsCount() == 0)
                {
                    MessageBox.Show("Выберите хотя бы один отдел, чтобы редактировать IMPORT_INFO.", "Family Manager", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                _importInfoBackup = record.IMPORT_INFO;     
                ShowRawJsonInEditor(record.IMPORT_INFO);        
                ImportInfoBox.IsReadOnly = false;      
                ImportInfoBox.Focusable = true;
                ImportInfoBox.Focus();

                ButtonEditFamilyInfo.Content = "💾";  
                ButtonEditFamilyInfo.ToolTip = "Сохранить IMPORT_INFO";

                _isEditingImportInfo = true;
                return;
            }

            try
            {
                var raw = GetRtbText(ImportInfoBox);
                var normalized = ValidateAndNormalizeJson(raw);  

                record.IMPORT_INFO = normalized;     

                ImportInfoBox.IsReadOnly = true;      
                ImportInfoBox.Focusable = true;           
                RenderImportInfo(record.IMPORT_INFO);   

                ButtonEditFamilyInfo.Content = "✏️";   
                ButtonEditFamilyInfo.ToolTip = "Редактировать IMPORT_INFO";

                _isEditingImportInfo = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("JSON невалиден:\n" + ex.Message,  "Проверка JSON", MessageBoxButton.OK, MessageBoxImage.Warning);

            }
        }









        // XAML. Запись данных в БД
        private void ButtonSaveDB_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Текущий record из DataContext
                var record = DataContext as FamilyManagerRecord;
                if (record == null)
                {
                    MessageBox.Show("Не удалось получить запись из контекста.", "Family Manager",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Собираем отделы из чекбоксов 
                var selectedDeps = DeptPanel.Children
                    .OfType<System.Windows.Controls.CheckBox>()
                    .Where(c => c.IsChecked == true)
                    .Select(c => c.Content?.ToString())
                    .Where(s => !string.IsNullOrWhiteSpace(s));

                record.DEPARTAMENT = JoinDepartments(selectedDeps);

                // STATUS
                var selectedUiStatus = StatusCombo.SelectedItem as string ?? "";
                var mappedStatus = MapStatusForSave(_originalStatus, selectedUiStatus);

                bool isIgnore = string.Equals(mappedStatus, "IGNORE", StringComparison.OrdinalIgnoreCase)
                     || string.Equals(_originalStatus, "IGNORE", StringComparison.OrdinalIgnoreCase);


                var selectedCategoryId = CategoryCombo.SelectedValue is int cid ? cid : 0; // CATEGORY (ID) и NC_NAME              
                var selectedProjectId = ProjectCombo.SelectedValue is int pid ? pid : 0; // PROJECT (ID)
                var selectedStageId = StageCombo.SelectedValue is int sid ? sid : 0; // STAGE (ID)
                var imgBytes = GetImageBytesFromImageControl(PreviewImage); // Копия IMAGE из PreviewImage

                if (!isIgnore)
                {
                    bool deptFilled = !string.IsNullOrWhiteSpace(record.DEPARTAMENT);
                    bool categoryOk = selectedCategoryId != 1;

                    List<string> errors = new List<string>();
                    if (!deptFilled) errors.Add("— Не выбран ни один отдел");
                    if (!categoryOk) errors.Add("— Не выбрана категория");

                    if (errors.Count > 0)
                    {
                        string msg = "Были допущены ошибки:\n" + string.Join("\n", errors);
                        MessageBox.Show(msg, "Проверка данных", MessageBoxButton.OK, MessageBoxImage.Information);
                        return; 
                    }
                }

                var result = new FamilyManagerRecord
                {
                    ID = record.ID,
                    STATUS = isIgnore ? mappedStatus : "OK",
                    FULLPATH = record.FULLPATH,
                    LM_DATE = record.LM_DATE,
                    CATEGORY = selectedCategoryId,
                    SUB_CATEGORY = record.SUB_CATEGORY,
                    PROJECT = selectedProjectId, 
                    STAGE = selectedStageId,   
                    DEPARTAMENT = record.DEPARTAMENT,
                    IMPORT_INFO = record.IMPORT_INFO,
                    IMAGE = imgBytes
                };

                ResultRecord = result;

                DeleteStatus = false;
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при сборе данных: " + ex.Message, "Family Manager",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // XAML. Удаление из БД
        private void ButtonDeleteDB_Click(object sender, RoutedEventArgs e)
        {
            DeleteStatus = true;
            DialogResult = true;
            Close();
        }
    }
}