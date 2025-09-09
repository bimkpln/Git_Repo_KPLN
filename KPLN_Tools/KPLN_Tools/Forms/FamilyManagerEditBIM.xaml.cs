using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media.Imaging;


namespace KPLN_Tools.Forms
{
    public partial class FamilyManagerEditBIM : Window
    {
        private const string DB_PATH = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_FamilyManager.db";

        private const int TARGET_MAX_SIDE = 512;   
        private const long JPEG_QUALITY = 85L; 
        private byte[] _preparedImageBytes;

        public FamilyManagerRecord ResultRecord { get; private set; }
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
            public int PROJECT { get; set; }
            public int STAGE { get; set; }
            public string DEPARTAMENT { get; set; }
            public string IMPORT_INFO { get; set; }
            public string CUSTOM_INFO { get; set; }
            public string INSTRUCTION_LINK { get; set; }
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

            CategoryCombo.ItemsSource = _categories;
            CategoryCombo.DisplayMemberPath = "NAME";
            CategoryCombo.SelectedValuePath = "ID";
            CategoryCombo.SelectedValue = record.CATEGORY;
            CategoryNcCombo.ItemsSource = _categories;
            CategoryNcCombo.DisplayMemberPath = "NC_NAME";
            CategoryNcCombo.SelectedValuePath = "ID";
            CategoryNcCombo.SelectedValue = record.CATEGORY;

            ProjectCombo.ItemsSource = _projects;
            ProjectCombo.DisplayMemberPath = "NAME";
            ProjectCombo.SelectedValuePath = "ID";
            ProjectCombo.SelectedValue = record.PROJECT;

            StageCombo.ItemsSource = _stages;
            StageCombo.DisplayMemberPath = "NAME";
            StageCombo.SelectedValuePath = "ID";
            StageCombo.SelectedValue = record.STAGE;
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
                        @"SELECT ID, STATUS, FULLPATH, LM_DATE, CATEGORY, PROJECT, STAGE,
                                 DEPARTAMENT, IMPORT_INFO, CUSTOM_INFO, INSTRUCTION_LINK, IMAGE
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
                            PROJECT = reader.IsDBNull(5) ? 0 : reader.GetInt32(5),
                            STAGE = reader.IsDBNull(6) ? 0 : reader.GetInt32(6),
                            DEPARTAMENT = reader.IsDBNull(7) ? "" : reader.GetString(7),
                            IMPORT_INFO = reader.IsDBNull(8) ? "" : reader.GetString(8),
                            CUSTOM_INFO = reader.IsDBNull(9) ? "" : reader.GetString(9),
                            INSTRUCTION_LINK = reader.IsDBNull(10) ? "" : reader.GetString(10),
                            IMAGE = reader.IsDBNull(11) ? null : (byte[])reader[11]
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
                case "UPDATE":
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
            // Если "Требует подтверждения", "Файл не найден..." или "В записи БД присутствует ошибка" — оставляем исходный
            if (selectedUiStatus == "Требует подтверждения" ||
                selectedUiStatus == "Файл не найден на диске" ||
                selectedUiStatus == "В записи БД присутствует ошибка")
            {
                return originalStatus;
            }

            if (selectedUiStatus == "Игнорировать в БД") return "IGNORE";
            if (selectedUiStatus == "Подтверждено") return "OK";

            // Иначе без изменений
            return originalStatus;
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
                            PROJECT = @project,
                            STAGE = @stage,
                            DEPARTAMENT = @dept,
                            IMPORT_INFO = @import,
                            CUSTOM_INFO = @custom,
                            INSTRUCTION_LINK = @link,
                            IMAGE = @image
                        WHERE ID = @id;";

                    cmd.Parameters.AddWithValue("@id", rec.ID);
                    cmd.Parameters.AddWithValue("@status", (object)(rec.STATUS ?? string.Empty));
                    cmd.Parameters.AddWithValue("@fullpath", (object)(rec.FULLPATH ?? string.Empty));
                    cmd.Parameters.AddWithValue("@lmdate", (object)(rec.LM_DATE ?? string.Empty));
                    cmd.Parameters.AddWithValue("@category", rec.CATEGORY);
                    cmd.Parameters.AddWithValue("@project", rec.PROJECT);
                    cmd.Parameters.AddWithValue("@stage", rec.STAGE);
                    cmd.Parameters.AddWithValue("@dept", (object)(rec.DEPARTAMENT ?? string.Empty));
                    cmd.Parameters.AddWithValue("@import", (object)(rec.IMPORT_INFO ?? string.Empty));
                    cmd.Parameters.AddWithValue("@custom", (object)(rec.CUSTOM_INFO ?? string.Empty));
                    cmd.Parameters.AddWithValue("@link", (object)(rec.INSTRUCTION_LINK ?? string.Empty));

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

                // STATUS
                var selectedUiStatus = StatusCombo.SelectedItem as string ?? "";
                var finalStatus = MapStatusForSave(_originalStatus, selectedUiStatus);

                // CATEGORY (ID) и NC_NAME
                var selectedCategoryId = CategoryCombo.SelectedValue is int cid ? cid : 0;
                var selectedCategory = _categories?.Find(c => c.ID == selectedCategoryId);
                var selectedCategoryNcName = selectedCategory?.NC_NAME ?? "";

                // PROJECT (ID)
                var selectedProjectId = ProjectCombo.SelectedValue is int pid ? pid : 0;

                // STAGE (ID)
                var selectedStageId = StageCombo.SelectedValue is int sid ? sid : 0;

                // Копия IMAGE из PreviewImage
                var imgBytes = GetImageBytesFromImageControl(PreviewImage);

                // Собираем финальный объект
                var result = new FamilyManagerRecord
                {
                    ID = record.ID,
                    STATUS = finalStatus,
                    FULLPATH = record.FULLPATH,
                    LM_DATE = record.LM_DATE,
                    CATEGORY = selectedCategoryId, 
                    PROJECT = selectedProjectId, 
                    STAGE = selectedStageId,   
                    DEPARTAMENT = record.DEPARTAMENT,
                    IMPORT_INFO = record.IMPORT_INFO,
                    CUSTOM_INFO = record.CUSTOM_INFO,
                    INSTRUCTION_LINK = record.INSTRUCTION_LINK,
                    IMAGE = imgBytes
                };

                ResultRecord = result;

                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при сборе данных: " + ex.Message, "Family Manager",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
