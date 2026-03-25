using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data.SQLite;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Drawing;

namespace KPLN_Tools.Forms
{
    public interface IApartmentManagerExternalController
    {
        void RequestPlaceApartment(int apartmentId);
        void RequestConvertTo3D(ApartmentPresetData presetData);
        void RequestOpenApartmentPresets(ApartmentPresetData presetData);
    }

    /// Данные пресета для дальнейших операций с квартирой. Используются при открытии преднастроек и при конвертации в 3D.
    public class ApartmentPresetData
    {
        public string SelectedPlanName { get; set; }

        public string LowerConstraint { get; set; }
        public string UpperConstraint { get; set; }

        public int BaseOffset { get; set; }
        public int WallHeight { get; set; }

        public Dictionary<int, string> WallTypeByThickness { get; set; }

        public string EntryDoor { get; set; }
        public string BathroomDoor { get; set; }
        public string RoomDoor { get; set; }

        public Dictionary<string, string> DoorsByRoomCategory { get; set; }

        /// Создание полной копии объекта, чтобы внешние вызовы не работали по той же ссылке.
        public ApartmentPresetData Clone()
        {
            return new ApartmentPresetData
            {
                SelectedPlanName = SelectedPlanName,
                LowerConstraint = LowerConstraint,
                UpperConstraint = UpperConstraint,
                BaseOffset = BaseOffset,
                WallHeight = WallHeight,
                WallTypeByThickness = WallTypeByThickness != null
                    ? new Dictionary<int, string>(WallTypeByThickness)
                    : new Dictionary<int, string>(),
                EntryDoor = EntryDoor,
                BathroomDoor = BathroomDoor,
                RoomDoor = RoomDoor,
                DoorsByRoomCategory = DoorsByRoomCategory != null
                    ? new Dictionary<string, string>(DoorsByRoomCategory)
                    : new Dictionary<string, string>()
            };
        }
    }

    public partial class ApartmentManagerWindow : Window
    {
        /// Номер отдела пользователя.
        public int _nDep;
        /// Внешний контроллер, который выполняет действия вне окна.
        private readonly IApartmentManagerExternalController _externalController;
        /// ID выбранной квартиры.
        public int SelectedApartmentId { get; private set; }
        /// Текущие данные пресета квартиры.
        public ApartmentPresetData ApartmentPresetData { get; private set; }
        /// Флаг запроса на конвертацию в 3D.
        public bool ConvertTo3DRequested { get; private set; }

        public ApartmentManagerWindow(
            int nDep,
            IApartmentManagerExternalController externalController,
            ApartmentPresetData presetData = null)
        {
            InitializeComponent();

            _nDep = nDep;
            _externalController = externalController;

            ApartmentPresetData = presetData != null
                ? presetData.Clone()
                : new ApartmentPresetData
                {
                    SelectedPlanName = "",
                    LowerConstraint = "",
                    UpperConstraint = "Неприсоединённая",
                    BaseOffset = 0,
                    WallHeight = 3000,
                    WallTypeByThickness = new Dictionary<int, string>(),
                    EntryDoor = "Не выбрано",
                    BathroomDoor = "Не выбрано",
                    RoomDoor = "Не выбрано",
                    DoorsByRoomCategory = new Dictionary<string, string>()
                };

            ApartmentManagerVm vm = new ApartmentManagerVm(_nDep);

            vm.ItemPicked += Vm_ItemPicked;
            vm.RequestClose += Vm_RequestClose;
            vm.ApartmentPresetsRequested += Vm_ApartmentPresetsRequested;
            vm.ConvertTo3DRequested += Vm_ConvertTo3DRequested;

            DataContext = vm;
        }

        /// Обновление текущих данных пресета извне. Хранится именно копия, а не исходная ссылка.
        public void SetApartmentPresetData(ApartmentPresetData data)
        {
            ApartmentPresetData = data != null ? data.Clone() : null;
        }

        /// Обработка выбора квартиры. Окно запоминает ID и передаёт запрос наружу.
        private void Vm_ItemPicked(int id)
        {
            SelectedApartmentId = id;

            if (_externalController == null)
                return;

            WindowState = WindowState.Minimized;
            _externalController.RequestPlaceApartment(id);
        }

        /// Запрос на открытие окна преднастроек квартир. Наружу уходит копия текущего пресета.
        private void Vm_ApartmentPresetsRequested()
        {
            if (_externalController == null)
                return;

            _externalController.RequestOpenApartmentPresets(
                ApartmentPresetData != null
                    ? ApartmentPresetData.Clone()
                    : null);
        }

        /// Запрос на конвертацию в 3D. Окно ставит флаг и передаёт действие наружу.
        private void Vm_ConvertTo3DRequested()
        {
            ConvertTo3DRequested = true;

            if (_externalController == null)
                return;

            WindowState = WindowState.Minimized;
            _externalController.RequestConvertTo3D(
                ApartmentPresetData != null
                    ? ApartmentPresetData.Clone()
                    : null);
        }

        /// Закрытие окна по запросу из VM.
        private void Vm_RequestClose()
        {
            Close();
        }
    }

    /// <summary>
    /// ViewModel главного окна.
    /// Здесь хранится список типов квартир, список элементов выбранного типа, команды интерфейса и вся логика работы с БД.
    /// </summary>
    internal class ApartmentManagerVm : INotifyPropertyChanged
    {
        /// Путь к SQLite-базе с квартирами.
        private const string DbPath = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_ApartmentManager.db";
        /// Папка, из которой берутся .rfa-файлы для синхронизации с БД.
        private const string RfaFolderPath = @"Z:\Отдел BIM\Туленинов Роман\aManager";

        /// Сигнал, что пользователь выбрал квартиру. Передаётся ID.
        public event Action<int> ItemPicked;
        /// Сигнал на открытие окна преднастроек.
        public event Action ApartmentPresetsRequested;
        /// Сигнал на запуск конвертации в 3D.
        public event Action ConvertTo3DRequested;
        /// Сигнал на закрытие окна.
        public event Action RequestClose;

        public event PropertyChangedEventHandler PropertyChanged;

        /// Признак "особого" режима для отдела BIM.
        public bool IsDep8 { get; private set; }

        /// Коллекция типов квартир для левого списка.
        public ObservableCollection<ApartmentTypeVm> ApartmentTypes { get; private set; }
        private ApartmentTypeVm _selectedType;

        /// Выбранный тип квартиры. При смене автоматически подгружается список элементов.
        public ApartmentTypeVm SelectedType
        {
            get { return _selectedType; }
            set
            {
                if (!ReferenceEquals(_selectedType, value))
                {
                    _selectedType = value;
                    OnPropertyChanged();
                    LoadItems();
                }
            }
        }

        /// Команда выбора карточки квартиры.
        public ICommand PickItemCommand { get; private set; }
        /// Команда загрузки/обновления картинки для квартиры.
        public ICommand UploadImageCommand { get; private set; }
        /// Команда открытия окна преднастроек.
        public ICommand OpenApartmentPresetsCommand { get; private set; }
        /// Команда запуска конвертации в 3D.
        public ICommand ConvertTo3DCommand { get; private set; }
        /// Команда закрытия окна.
        public ICommand CloseCommand { get; private set; }

        /// <summary>
        /// Команда синхронизации БД с .rfa-файлами.
        /// </summary>
        public ICommand UpdateDbCommand { get; private set; }

        public ApartmentManagerVm(int nDep)
        {
            IsDep8 = (nDep == 8);

            ApartmentTypes = new ObservableCollection<ApartmentTypeVm>();

            PickItemCommand = new RelayCommand<ApartmentItemVm>(OnPick);
            CloseCommand = new RelayCommand(OnClose);
            UploadImageCommand = new RelayCommand<ApartmentItemVm>(OnUploadImage);
            OpenApartmentPresetsCommand = new RelayCommand(OnOpenApartmentPresets);
            ConvertTo3DCommand = new RelayCommand(OnConvertTo3D);
            UpdateDbCommand = new RelayCommand<Window>(OnUpdateDb);

            LoadTypes();
        }

        /// Проброс события открытия преднастроек наружу.
        private void OnOpenApartmentPresets()  { ApartmentPresetsRequested?.Invoke(); }

        /// Проброс события конвертации в 3D наружу.
        private void OnConvertTo3D()
        {
            ConvertTo3DRequested?.Invoke();
        }

        /// Загружает из БД список уникальных типов квартир.
        private void LoadTypes()
        {
            ApartmentTypes.Clear();

            if (!File.Exists(DbPath))
            {
                MessageBox.Show("Не найдена база:\n" + DbPath, "ApartmentManager");
                return;
            }

            try
            {
                using (var con = OpenConnection(DbPath, true))
                using (var cmd = con.CreateCommand())
                {
                    cmd.CommandText =
                        "SELECT DISTINCT TRIM(ATYPE) AS ATYPE " +
                        "FROM Main " +
                        "WHERE ATYPE IS NOT NULL AND TRIM(ATYPE) <> '' " +
                        "ORDER BY TRIM(ATYPE);";

                    using (var r = cmd.ExecuteReader())
                    {
                        while (r.Read())
                        {
                            string atypeRaw = r.IsDBNull(0) ? "" : r.GetString(0);
                            string atype = NormalizeAtype(atypeRaw);

                            if (!string.IsNullOrWhiteSpace(atype))
                                ApartmentTypes.Add(new ApartmentTypeVm { Name = atype });
                        }
                    }
                }

                if (ApartmentTypes.Count > 0)
                    SelectedType = ApartmentTypes[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка чтения типов:\n" + ex, "ApartmentManager");
            }
        }

        /// Загружает планировки для выбранного типа квартиры. Использует поле PIC как источник изображения-превью.
        private void LoadItems()
        {
            if (SelectedType == null)
                return;

            SelectedType.Items.Clear();

            if (!File.Exists(DbPath))
                return;

            string atype = NormalizeAtype(SelectedType.Name);

            try
            {
                using (var con = OpenConnection(DbPath, true))
                using (var cmd = con.CreateCommand())
                {
                    cmd.CommandText =
                        "SELECT ID, PIC " +
                        "FROM Main " +
                        "WHERE TRIM(REPLACE(REPLACE(ATYPE,'K','К'),'k','К')) = @atype " +
                        "ORDER BY ID;";

                    cmd.Parameters.AddWithValue("@atype", atype);

                    using (var r = cmd.ExecuteReader())
                    {
                        while (r.Read())
                        {
                            int id = r.GetInt32(0);
                            byte[] bytes = null;

                            if (!r.IsDBNull(1))
                                bytes = (byte[])r["PIC"];

                            SelectedType.Items.Add(new ApartmentItemVm
                            {
                                Id = id,
                                Title = SelectedType.Name + " #" + id,
                                Preview = BytesToImage(bytes)
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка чтения планировок:\n" + ex, "ApartmentManager");
            }
        }

        /// Обновление изображения для выбранной квартиры. Доступно только для отдела BIM.
        private void OnUploadImage(ApartmentItemVm item)
        {
            if (item == null || !IsDep8)
                return;

            var ofd = new OpenFileDialog
            {
                Title = "Выберите изображение планировки",
                Filter = "Изображения (*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg|Все файлы (*.*)|*.*",
                Multiselect = false
            };

            bool? ok = ofd.ShowDialog();
            if (ok != true)
                return;

            string filePath = ofd.FileName;
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                return;

            try
            {
                byte[] bytes = File.ReadAllBytes(filePath);

                using (var con = OpenConnection(DbPath, false))
                using (var cmd = con.CreateCommand())
                {
                    cmd.CommandText = "UPDATE Main SET PIC = @pic WHERE ID = @id;";
                    cmd.Parameters.AddWithValue("@id", item.Id);

                    var p = cmd.CreateParameter();
                    p.ParameterName = "@pic";
                    p.DbType = System.Data.DbType.Binary;
                    p.Value = bytes;
                    cmd.Parameters.Add(p);

                    int affected = cmd.ExecuteNonQuery();
                    if (affected <= 0)
                    {
                        MessageBox.Show("Не удалось обновить PIC для ID=" + item.Id, "ApartmentManager");
                        return;
                    }
                }

                item.Preview = BytesToImage(bytes);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка загрузки/сохранения изображения:\n" + ex, "ApartmentManager");
            }
        }

        /// Преобразует массив байт в WPF ImageSource. Если картинка битая или не читается — возвращает null.
        private static ImageSource BytesToImage(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0)
                return null;

            try
            {
                using (var ms = new MemoryStream(bytes))
                {
                    var img = new BitmapImage();
                    img.BeginInit();
                    img.CacheOption = BitmapCacheOption.OnLoad;
                    img.StreamSource = ms;
                    img.EndInit();
                    img.Freeze();
                    return img;
                }
            }
            catch
            {
                return null;
            }
        }

        /// Открывает SQLite-соединение в режиме чтения или чтения/записи.
        private static SQLiteConnection OpenConnection(string dbPath, bool readOnly)
        {
            string cs = readOnly
                ? "Data Source=" + dbPath + ";Version=3;Read Only=True;"
                : "Data Source=" + dbPath + ";Version=3;";

            var con = new SQLiteConnection(cs);
            con.Open();
            return con;
        }

        /// Синхронизирует БД с папкой .rfa-файлов:
        private void OnUpdateDb(Window ownerWindow)
        {
            if (!IsDep8)
                return;

            if (!File.Exists(DbPath))
            {
                MessageBox.Show("Не найдена база:\n" + DbPath, "ApartmentManager");
                return;
            }

            if (!Directory.Exists(RfaFolderPath))
            {
                MessageBox.Show("Не найдена папка:\n" + RfaFolderPath, "ApartmentManager");
                return;
            }

            try
            {
                string[] files = Directory.GetFiles(RfaFolderPath, "*.rfa", SearchOption.TopDirectoryOnly);
                HashSet<string> actualPaths = new HashSet<string>(files, StringComparer.OrdinalIgnoreCase);

                Dictionary<string, int> dbItemsByPath = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                using (var con = OpenConnection(DbPath, true))
                using (var cmd = con.CreateCommand())
                {
                    cmd.CommandText =
                        "SELECT ID, FPATH " +
                        "FROM Main " +
                        "WHERE FPATH IS NOT NULL AND TRIM(FPATH) <> '';";

                    using (var r = cmd.ExecuteReader())
                    {
                        while (r.Read())
                        {
                            int id = r.GetInt32(0);
                            string path = r.IsDBNull(1) ? null : r.GetString(1);

                            if (string.IsNullOrWhiteSpace(path))
                                continue;

                            path = path.Trim();

                            if (!dbItemsByPath.ContainsKey(path))
                                dbItemsByPath.Add(path, id);
                        }
                    }
                }

                List<KeyValuePair<string, int>> itemsToDelete = dbItemsByPath
                    .Where(x => !actualPaths.Contains(x.Key))
                    .ToList();

                List<string> deletedNames = new List<string>();

                if (itemsToDelete.Count > 0)
                {
                    using (var con = OpenConnection(DbPath, false))
                    using (var tx = con.BeginTransaction())
                    {
                        foreach (var kvp in itemsToDelete)
                        {
                            using (var deleteCmd = con.CreateCommand())
                            {
                                deleteCmd.Transaction = tx;
                                deleteCmd.CommandText = "DELETE FROM Main WHERE ID = @id;";
                                deleteCmd.Parameters.AddWithValue("@id", kvp.Value);
                                deleteCmd.ExecuteNonQuery();
                            }

                            deletedNames.Add(Path.GetFileNameWithoutExtension(kvp.Key));
                        }

                        tx.Commit();
                    }
                }

                List<ApartmentImportItemVm> newItems = new List<ApartmentImportItemVm>();

                foreach (string file in files)
                {
                    if (dbItemsByPath.ContainsKey(file))
                        continue;

                    newItems.Add(new ApartmentImportItemVm
                    {
                        FilePath = file,
                        FileName = Path.GetFileNameWithoutExtension(file),
                        Preview = ShellPreviewHelper.GetShellPreviewImage(file)
                    });
                }

                int addedCount = 0;

                if (newItems.Count > 0)
                {
                    var wnd = new ApartmentImportWindow(newItems);
                    if (ownerWindow != null)
                        wnd.Owner = ownerWindow;

                    bool? res = wnd.ShowDialog();
                    if (res == true)
                    {
                        List<ApartmentImportItemVm> itemsToInsert = wnd.Items
                            .Where(x => !string.IsNullOrWhiteSpace(x.SelectedAtype))
                            .ToList();

                        if (itemsToInsert.Count > 0)
                        {
                            using (var con = OpenConnection(DbPath, false))
                            using (var tx = con.BeginTransaction())
                            {
                                int nextId;

                                using (var getMaxCmd = con.CreateCommand())
                                {
                                    getMaxCmd.Transaction = tx;
                                    getMaxCmd.CommandText = "SELECT IFNULL(MAX(ID), 0) FROM Main;";
                                    object o = getMaxCmd.ExecuteScalar();
                                    nextId = Convert.ToInt32(o) + 1;
                                }

                                foreach (var item in itemsToInsert)
                                {
                                    byte[] picBytes = ShellPreviewHelper.GetShellPreviewBytes(item.FilePath);

                                    using (var insertCmd = con.CreateCommand())
                                    {
                                        insertCmd.Transaction = tx;
                                        insertCmd.CommandText =
                                            "INSERT INTO Main (ID, FPATH, VNAME, ATYPE, PIC) " +
                                            "VALUES (@id, @fpath, @vname, @atype, @pic);";

                                        insertCmd.Parameters.AddWithValue("@id", nextId++);
                                        insertCmd.Parameters.AddWithValue("@fpath", item.FilePath);
                                        insertCmd.Parameters.AddWithValue("@vname", item.FileName);
                                        insertCmd.Parameters.AddWithValue("@atype", item.SelectedAtype);

                                        var pPic = insertCmd.CreateParameter();
                                        pPic.ParameterName = "@pic";
                                        pPic.DbType = System.Data.DbType.Binary;
                                        pPic.Value = (object)picBytes ?? DBNull.Value;
                                        insertCmd.Parameters.Add(pPic);

                                        insertCmd.ExecuteNonQuery();
                                    }
                                }

                                tx.Commit();
                            }

                            addedCount = itemsToInsert.Count;
                        }
                    }
                }

                LoadTypes();

                if (deletedNames.Count == 0 && addedCount == 0)
                {
                    MessageBox.Show("БД уже актуальна. Удалённых и новых файлов не найдено.", "ApartmentManager");
                    return;
                }

                string message = "";

                if (deletedNames.Count > 0)
                {
                    message += "Удалено из БД: " + deletedNames.Count;

                    if (deletedNames.Count <= 20)
                        message += "\n" + string.Join("\n", deletedNames);
                }

                if (addedCount > 0)
                {
                    if (!string.IsNullOrWhiteSpace(message))
                        message += "\n\n";

                    message += "Добавлено в БД: " + addedCount;
                }

                MessageBox.Show(message, "ApartmentManager");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка обновления БД:\n" + ex, "ApartmentManager");
            }
        }

        /// Обработка выбора карточки квартиры.
        private void OnPick(ApartmentItemVm item)
        {
            if (item == null)
                return;

            ItemPicked?.Invoke(item.Id);
        }

        /// Нормализует обозначение типа квартиры
        private static string NormalizeAtype(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
                return "";

            s = s.Trim();
            s = s.Replace('K', 'К');
            s = s.Replace('k', 'К');
            return s;
        }

        private void OnPropertyChanged([CallerMemberName] string p = null)
        {
            var h = PropertyChanged;
            if (h != null)
                h(this, new PropertyChangedEventArgs(p));
        }

        /// Обработка закрытия.
        private void OnClose()
        {
            RequestClose?.Invoke();
        }
    }




    /// <summary>
    /// VM типа квартиры для левого списка. Содержит имя типа и коллекцию его элементов.
    /// </summary>
    internal class ApartmentTypeVm
    {
        public string Name { get; set; }

        public ObservableCollection<ApartmentItemVm> Items { get; private set; }

        public ApartmentTypeVm()
        {
            Items = new ObservableCollection<ApartmentItemVm>();
        }
    }

    /// <summary>
    /// VM карточки квартиры. Содержит ID, заголовок и превью-картинку.
    /// </summary>
    internal class ApartmentItemVm : INotifyPropertyChanged
    {
        public int Id { get; set; }
        public string Title { get; set; }

        private ImageSource _preview;

        public ImageSource Preview
        {
            get { return _preview; }
            set
            {
                _preview = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string p = null)
        {
            var h = PropertyChanged;
            if (h != null)
                h(this, new PropertyChangedEventArgs(p));
        }
    }

    /// <summary>
    /// VM элемента, который предлагается импортировать в БД. Хранит путь к файлу, имя, выбранный тип и превью.
    /// </summary>
    public class ApartmentImportItemVm : INotifyPropertyChanged
    {
        public string FilePath { get; set; }
        public string FileName { get; set; }

        private string _selectedAtype;
        public string SelectedAtype
        {
            get { return _selectedAtype; }
            set
            {
                _selectedAtype = value;
                OnPropertyChanged();
            }
        }

        private ImageSource _preview;
        public ImageSource Preview
        {
            get { return _preview; }
            set
            {
                _preview = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string p = null)
        {
            var h = PropertyChanged;
            if (h != null)
                h(this, new PropertyChangedEventArgs(p));
        }
    }

    /// <summary>
    /// Простая команда без параметра. Всегда доступна для выполнения.
    /// </summary>
    internal class RelayCommand : ICommand
    {
        private readonly Action _execute;

        public RelayCommand(Action e)
        {
            _execute = e;
        }

        public bool CanExecute(object p)
        {
            return true;
        }

        public void Execute(object p)
        {
            _execute();
        }

        public event EventHandler CanExecuteChanged;
    }

    /// <summary>
    /// Простая команда с параметром. В текущей реализации, если параметр null, команда просто ничего не делает.
    /// </summary>
    internal class RelayCommand<T> : ICommand
    {
        private readonly Action<T> _execute;

        public RelayCommand(Action<T> e)
        {
            _execute = e;
        }

        public bool CanExecute(object p)
        {
            return true;
        }

        public void Execute(object p)
        {
            if (p == null)
                return;

            _execute((T)p);
        }

        public event EventHandler CanExecuteChanged;
    }

    /// Конвертер для управления видимостью по null/not null. Нужен в XAML для показа картинки или текстовой заглушки.
    public class NullToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool invert = parameter != null &&
                          string.Equals(parameter.ToString(), "NotNull", StringComparison.OrdinalIgnoreCase);

            bool isNull = (value == null);

            if (!invert)
                return isNull ? Visibility.Visible : Visibility.Collapsed;

            return isNull ? Visibility.Collapsed : Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }

    /// <summary>
    /// Вспомогательный класс для получения shell-preview и иконок файлов. Используется при импорте новых .rfa-элементов в БД.
    /// </summary>
    internal static class ShellPreviewHelper
    {
        /// Получает превью как ImageSource для показа в WPF.
        public static ImageSource GetShellPreviewImage(string filePath)
        {
            try
            {
                byte[] bytes = GetShellPreviewBytes(filePath);
                if (bytes == null || bytes.Length == 0)
                    return null;

                using (var ms = new MemoryStream(bytes))
                {
                    var img = new BitmapImage();
                    img.BeginInit();
                    img.CacheOption = BitmapCacheOption.OnLoad;
                    img.StreamSource = ms;
                    img.EndInit();
                    img.Freeze();
                    return img;
                }
            }
            catch
            {
                return null;
            }
        }

        /// Получает превью файла в виде массива байт PNG. Сначала пытается взять shell-thumbnail, если не получилось — берёт ассоциированную иконку файла.
        public static byte[] GetShellPreviewBytes(string filePath)
        {
            Bitmap bmp = null;

            try
            {
                bmp = GetShellThumbnail(filePath, 512);

                if (bmp == null)
                    bmp = GetAssociatedIconBitmap(filePath);

                if (bmp == null)
                    return null;

                using (bmp)
                using (var ms = new MemoryStream())
                {
                    bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                    return ms.ToArray();
                }
            }
            catch
            {
                return null;
            }
        }

        /// Возвращает bitmap ассоциированной системной иконки файла.
        private static Bitmap GetAssociatedIconBitmap(string filePath)
        {
            try
            {
                Icon icon = Icon.ExtractAssociatedIcon(filePath);
                if (icon == null)
                    return null;

                using (icon)
                {
                    return icon.ToBitmap();
                }
            }
            catch
            {
                return null;
            }
        }

        /// Получает thumbnail через Windows Shell API.
        private static Bitmap GetShellThumbnail(string filePath, int size)
        {
            IShellItemImageFactory factory = null;
            IntPtr hBitmap = IntPtr.Zero;

            try
            {
                Guid shellItemGuid = new Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE");
                SHCreateItemFromParsingName(filePath, IntPtr.Zero, ref shellItemGuid, out factory);

                SIZE s;
                s.cx = size;
                s.cy = size;

                factory.GetImage(
                    s,
                    SIIGBF.BIGGERSIZEOK | SIIGBF.THUMBNAILONLY,
                    out hBitmap);

                if (hBitmap == IntPtr.Zero)
                    return null;

                using (Bitmap temp = System.Drawing.Image.FromHbitmap(hBitmap))
                {
                    return new Bitmap(temp);
                }
            }
            catch
            {
                return null;
            }
            finally
            {
                if (hBitmap != IntPtr.Zero)
                    DeleteObject(hBitmap);

                if (factory != null)
                    Marshal.ReleaseComObject(factory);
            }
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void SHCreateItemFromParsingName(
            [MarshalAs(UnmanagedType.LPWStr)] string pszPath,
            IntPtr pbc,
            ref Guid riid,
            [MarshalAs(UnmanagedType.Interface)] out IShellItemImageFactory ppv);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        /// COM-интерфейс Windows Shell для получения изображения элемента.
        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("BCC18B79-BA16-442F-80C4-8A59C30C463B")]
        private interface IShellItemImageFactory
        {
            void GetImage(SIZE size, SIIGBF flags, out IntPtr phbm);
        }

        /// Структура размера для Shell API.
        [StructLayout(LayoutKind.Sequential)]
        private struct SIZE
        {
            public int cx;
            public int cy;
        }

        /// Флаги получения изображения из Windows Shell.
        [Flags]
        private enum SIIGBF
        {
            RESIZETOFIT = 0x00,
            BIGGERSIZEOK = 0x01,
            MEMORYONLY = 0x02,
            ICONONLY = 0x04,
            THUMBNAILONLY = 0x08,
            INCACHEONLY = 0x10
        }
    }
}