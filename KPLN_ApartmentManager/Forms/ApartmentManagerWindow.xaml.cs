using KPLN_ApartmentManager.Forms;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace KPLN_ApartmentManager.Forms
{
    public interface IApartmentManagerExternalController
    {
        void RequestPlaceApartment(int apartmentId);
        void RequestConvertTo3D(ApartmentPresetData presetData);
        void RequestRefreshApartmentPresets(ApartmentPresetData presetData);
        void RequestUpdateApartmentMarks();
    }

    public enum ApartmentFamilyPostProcessAction
    {
        Save2DFamiliesFromUnderlay,
        Keep2DUnderlay,
        Delete2DUnderlay
    }

    public class ApartmentFamilyPostProcessActionOption
    {
        public ApartmentFamilyPostProcessAction Value { get; set; }
        public string Title { get; set; }
    }

    public class ApartmentPresetData
    {
        public string SelectedPlanName { get; set; }
        public string SelectedPlanModelSignature { get; set; }

        public string LowerConstraint { get; set; }
        public string UpperConstraint { get; set; }

        public double BaseOffset { get; set; }
        public double WallHeight { get; set; }

        public Dictionary<int, string> WallTypeByThickness { get; set; }

        public string WindowType { get; set; }
        public int WindowSillHeight { get; set; }

        public string EntryDoor { get; set; }
        public string BathroomDoor { get; set; }
        public string RoomDoor { get; set; }

        public Dictionary<string, string> DoorsByRoomCategory { get; set; }

        public ApartmentFamilyPostProcessAction FamilyPostProcessAction { get; set; }

        public ApartmentPresetData Clone()
        {
            return new ApartmentPresetData
            {
                SelectedPlanName = SelectedPlanName,
                SelectedPlanModelSignature = SelectedPlanModelSignature,
                LowerConstraint = LowerConstraint,
                UpperConstraint = UpperConstraint,
                BaseOffset = BaseOffset,
                WallHeight = WallHeight,
                WallTypeByThickness = WallTypeByThickness != null
                    ? new Dictionary<int, string>(WallTypeByThickness)
                    : new Dictionary<int, string>(),
                WindowType = WindowType,
                WindowSillHeight = WindowSillHeight,
                EntryDoor = EntryDoor,
                BathroomDoor = BathroomDoor,
                RoomDoor = RoomDoor,
                DoorsByRoomCategory = DoorsByRoomCategory != null
                    ? new Dictionary<string, string>(DoorsByRoomCategory)
                    : new Dictionary<string, string>(),
                FamilyPostProcessAction = FamilyPostProcessAction
            };
        }
    }

    public partial class ApartmentManagerWindow : Window
    {
        public int _nDep;
        private readonly IApartmentManagerExternalController _externalController;
        private readonly ApartmentManagerVm _vm;
        public int SelectedApartmentId { get; private set; }
        public ApartmentPresetData ApartmentPresetData { get; private set; }
        public bool ConvertTo3DRequested { get; private set; }

        public ApartmentManagerWindow(
            int nDep,
            IApartmentManagerExternalController externalController,
            ApartmentPresetData presetData = null,
            ApartmentPresetPanelContext presetContext = null)
        {
            InitializeComponent();

            _nDep = nDep;
            _externalController = externalController;

            ApartmentPresetData = presetData != null
                ? presetData.Clone()
                : new ApartmentPresetData
                {
                    SelectedPlanName = "",
                    SelectedPlanModelSignature = "",
                    LowerConstraint = "",
                    UpperConstraint = "Неприсоединённая",
                    BaseOffset = 0,
                    WallHeight = 3000,
                    WallTypeByThickness = new Dictionary<int, string>(),
                    WindowType = "Не выбрано",
                    WindowSillHeight = 900,
                    EntryDoor = "Не выбрано",
                    BathroomDoor = "Не выбрано",
                    RoomDoor = "Не выбрано",
                    DoorsByRoomCategory = new Dictionary<string, string>(),
                    FamilyPostProcessAction = ApartmentFamilyPostProcessAction.Save2DFamiliesFromUnderlay
                };

            _vm = new ApartmentManagerVm(_nDep, ApartmentPresetData, presetContext);

            _vm.ItemPicked += Vm_ItemPicked;
            _vm.RequestClose += Vm_RequestClose;
            _vm.ApartmentPresetDataRefreshRequested += Vm_ApartmentPresetDataRefreshRequested;
            _vm.ConvertTo3DRequested += Vm_ConvertTo3DRequested;
            _vm.FamilyPostProcessActionChanged += Vm_FamilyPostProcessActionChanged;
            _vm.ApartmentMarksUpdateRequested += Vm_ApartmentMarksUpdateRequested;
            _vm.ApartmentPresetDataChanged += Vm_ApartmentPresetDataChanged;

            Closing += ApartmentManagerWindow_Closing;

            DataContext = _vm;
        }

        public void SetApartmentPresetData(ApartmentPresetData data)
        {
            ApartmentPresetData = data != null ? data.Clone() : null;
            if (_vm != null)
                _vm.SetApartmentPresetData(ApartmentPresetData);
        }

        public void SetApartmentPresetContext(ApartmentPresetPanelContext context, ApartmentPresetData currentData)
        {
            if (_vm == null)
                return;

            _vm.SetApartmentPresetContext(context, currentData);
            SyncApartmentPresetDataFromVm();
        }

        private void Vm_ItemPicked(int id)
        {
            SelectedApartmentId = id;

            if (_externalController == null)
                return;

            WindowState = WindowState.Minimized;
            _externalController.RequestPlaceApartment(id);
        }

        private void Vm_FamilyPostProcessActionChanged(ApartmentFamilyPostProcessAction action)
        {
            if (ApartmentPresetData == null)
                ApartmentPresetData = new ApartmentPresetData();

            ApartmentPresetData.FamilyPostProcessAction = action;
        }

        private void Vm_ApartmentPresetDataRefreshRequested()
        {
            if (_externalController == null)
                return;

            SyncApartmentPresetDataFromVm();

            _externalController.RequestRefreshApartmentPresets(
                ApartmentPresetData != null
                    ? ApartmentPresetData.Clone()
                    : null);
        }

        private void Vm_ConvertTo3DRequested()
        {
            ConvertTo3DRequested = true;

            if (_externalController == null)
                return;

            SyncApartmentPresetDataFromVm();

            WindowState = WindowState.Minimized;
            _externalController.RequestConvertTo3D(
                ApartmentPresetData != null
                    ? ApartmentPresetData.Clone()
                    : null);
        }

        private void Vm_ApartmentMarksUpdateRequested()
        {
            if (_externalController == null)
                return;

            WindowState = WindowState.Minimized;
            _externalController.RequestUpdateApartmentMarks();
        }

        private void Vm_RequestClose()
        {
            Close();
        }

        private void Vm_ApartmentPresetDataChanged(ApartmentPresetData data)
        {
            ApartmentPresetData = data != null ? data.Clone() : null;
        }

        private void ApartmentManagerWindow_Closing(object sender, CancelEventArgs e)
        {
            SyncApartmentPresetDataFromVm();
        }

        public void MarkApartmentPresetDataStale()
        {
            if (_vm != null)
                _vm.MarkApartmentPresetDataStale();
        }

        private void SyncApartmentPresetDataFromVm()
        {
            if (_vm == null)
                return;

            ApartmentPresetData = _vm.GetApartmentPresetData();
        }

        public ApartmentPresetPanelContext GetApartmentPresetContextSnapshot()
        {
            return _vm != null
                ? _vm.GetApartmentPresetContextSnapshot()
                : new ApartmentPresetPanelContext();
        }

        private void DigitsOnly_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            e.Handled = e.Text.Any(x => !char.IsDigit(x));
        }

        private void DigitsOnly_Pasting(object sender, DataObjectPastingEventArgs e)
        {
            if (!e.DataObject.GetDataPresent(typeof(string)))
            {
                e.CancelCommand();
                return;
            }

            string text = e.DataObject.GetData(typeof(string)) as string;
            if (string.IsNullOrWhiteSpace(text) || text.Any(x => !char.IsDigit(x)))
                e.CancelCommand();
        }

        private void SignedDecimal_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            if (textBox == null)
            {
                e.Handled = true;
                return;
            }

            e.Handled = !IsPotentialDecimalText(GetTextAfterInput(textBox, e.Text));
        }

        private void SignedDecimal_Pasting(object sender, DataObjectPastingEventArgs e)
        {
            if (!e.DataObject.GetDataPresent(typeof(string)))
            {
                e.CancelCommand();
                return;
            }

            TextBox textBox = sender as TextBox;
            string text = e.DataObject.GetData(typeof(string)) as string;
            string nextText = textBox != null
                ? GetTextAfterInput(textBox, text)
                : text;

            if (!IsPotentialDecimalText(nextText))
                e.CancelCommand();
        }

        private void PresetComboBox_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            ComboBox comboBox = sender as ComboBox;
            if (comboBox == null || comboBox.IsDropDownOpen)
                return;

            e.Handled = true;

            UIElement parent = comboBox.Parent as UIElement;
            if (parent == null)
                return;

            MouseWheelEventArgs args = new MouseWheelEventArgs(e.MouseDevice, e.Timestamp, e.Delta)
            {
                RoutedEvent = UIElement.MouseWheelEvent,
                Source = sender
            };
            parent.RaiseEvent(args);
        }

        private static string GetTextAfterInput(TextBox textBox, string input)
        {
            string text = textBox.Text ?? "";
            input = input ?? "";

            int selectionStart = textBox.SelectionStart;
            if (selectionStart < 0)
                selectionStart = 0;
            if (selectionStart > text.Length)
                selectionStart = text.Length;

            int selectionLength = textBox.SelectionLength;
            if (selectionLength < 0)
                selectionLength = 0;
            if (selectionStart + selectionLength > text.Length)
                selectionLength = text.Length - selectionStart;

            return text.Remove(selectionStart, selectionLength).Insert(selectionStart, input);
        }

        private static bool IsPotentialDecimalText(string text)
        {
            if (string.IsNullOrEmpty(text))
                return true;

            bool hasSeparator = false;

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];

                if (char.IsDigit(c))
                    continue;

                if (c == '-')
                {
                    if (i != 0)
                        return false;
                    continue;
                }

                if (c == '.' || c == ',')
                {
                    if (hasSeparator)
                        return false;

                    hasSeparator = true;
                    continue;
                }

                return false;
            }

            return true;
        }
    }

    internal class ApartmentManagerVm : INotifyPropertyChanged
    {
        private const string DbPath = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_ApartmentManager.db";
        private const string RfaFolderPath = @"X:\BIM\3_Семейства\1_АР\000_Архитектурная концепция\000_Семейства квартир";
        private const int Long3DConversionApartmentWarningThreshold = 10;

        public event Action<int> ItemPicked;
        public event Action ApartmentPresetDataRefreshRequested;
        public event Action ConvertTo3DRequested;
        public event Action ApartmentMarksUpdateRequested;
        public event Action<ApartmentPresetData> ApartmentPresetDataChanged;
        public event Action<ApartmentFamilyPostProcessAction> FamilyPostProcessActionChanged;
        public event Action RequestClose;

        public event PropertyChangedEventHandler PropertyChanged;

        public bool IsDep8 { get; private set; }

        public ObservableCollection<ApartmentTypeVm> ApartmentTypes { get; private set; }

        public ObservableCollection<ApartmentFamilyPostProcessActionOption> FamilyPostProcessActions { get; private set; }

        public ApartmentPresetsVm PresetsVm { get; private set; }

        private ApartmentFamilyPostProcessActionOption _selectedFamilyPostProcessAction;
        public ApartmentFamilyPostProcessActionOption SelectedFamilyPostProcessAction
        {
            get { return _selectedFamilyPostProcessAction; }
            set
            {
                if (!ReferenceEquals(_selectedFamilyPostProcessAction, value))
                {
                    _selectedFamilyPostProcessAction = value;
                    OnPropertyChanged();

                    if (_selectedFamilyPostProcessAction != null)
                    {
                        if (PresetsVm != null)
                            PresetsVm.SetFamilyPostProcessAction(_selectedFamilyPostProcessAction.Value);

                        FamilyPostProcessActionChanged?.Invoke(_selectedFamilyPostProcessAction.Value);
                    }
                }
            }
        }

        private ApartmentTypeVm _selectedType;

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

        public ICommand PickItemCommand { get; private set; }
        public ICommand UploadImageCommand { get; private set; }
        public ICommand OpenApartmentPresetsCommand { get; private set; }
        public ICommand RefreshApartmentPresetsCommand { get; private set; }
        public ICommand ConvertTo3DCommand { get; private set; }
        public ICommand UpdateApartmentMarksCommand { get; private set; }
        public ICommand CloseCommand { get; private set; }
        public ICommand UpdateDbCommand { get; private set; }

        public ApartmentManagerVm(int nDep, ApartmentPresetData initialPresetData, ApartmentPresetPanelContext initialPresetContext)
        {
            IsDep8 = (nDep == 8);

            ApartmentTypes = new ObservableCollection<ApartmentTypeVm>();
            PresetsVm = new ApartmentPresetsVm(initialPresetData, initialPresetContext ?? new ApartmentPresetPanelContext());
            PresetsVm.DataChanged += PresetsVm_DataChanged;
            PresetsVm.StateChanged += PresetsVm_StateChanged;

            FamilyPostProcessActions = new ObservableCollection<ApartmentFamilyPostProcessActionOption>
            {
                new ApartmentFamilyPostProcessActionOption
                {
                    Value = ApartmentFamilyPostProcessAction.Save2DFamiliesFromUnderlay,
                    Title = "Сохранить 2D-семейства с подложки"
                },
                new ApartmentFamilyPostProcessActionOption
                {
                    Value = ApartmentFamilyPostProcessAction.Keep2DUnderlay,
                    Title = "Сохранить 2D-подложку"
                },
                new ApartmentFamilyPostProcessActionOption
                {
                    Value = ApartmentFamilyPostProcessAction.Delete2DUnderlay,
                    Title = "Полностью удалить 2D-подложку"
                }
            };

            PickItemCommand = new RelayCommand<ApartmentItemVm>(OnPick);
            CloseCommand = new RelayCommand(OnClose);
            UploadImageCommand = new RelayCommand<ApartmentItemVm>(OnUploadImage);
            OpenApartmentPresetsCommand = new RelayCommand(OnOpenApartmentPresets);
            RefreshApartmentPresetsCommand = new RelayCommand(OnOpenApartmentPresets);
            ConvertTo3DCommand = new RelayCommand(OnConvertTo3D, CanConvertTo3D);
            UpdateApartmentMarksCommand = new RelayCommand(OnUpdateApartmentMarks);
            UpdateDbCommand = new RelayCommand<Window>(OnUpdateDb);

            SelectedFamilyPostProcessAction = FamilyPostProcessActions
                .FirstOrDefault(x => initialPresetData != null && x.Value == initialPresetData.FamilyPostProcessAction)
                ?? FamilyPostProcessActions.FirstOrDefault();

            LoadTypes();
        }

        public ApartmentPresetData GetApartmentPresetData()
        {
            return PresetsVm != null ? PresetsVm.BuildData() : new ApartmentPresetData();
        }

        public ApartmentPresetPanelContext GetApartmentPresetContextSnapshot()
        {
            return PresetsVm != null
                ? PresetsVm.GetContextSnapshot()
                : new ApartmentPresetPanelContext();
        }

        public void SetApartmentPresetData(ApartmentPresetData data)
        {
            if (PresetsVm != null)
                PresetsVm.ApplyContext(new ApartmentPresetPanelContext(), data);
        }

        public void SetApartmentPresetContext(ApartmentPresetPanelContext context, ApartmentPresetData currentData)
        {
            if (PresetsVm != null)
                PresetsVm.ApplyContext(context, currentData);
        }

        public void MarkApartmentPresetDataStale()
        {
            if (PresetsVm != null)
                PresetsVm.MarkDataStale();
        }

        private void OnOpenApartmentPresets()
        {
            ApartmentPresetDataRefreshRequested?.Invoke();
        }

        private bool CanConvertTo3D()
        {
            return PresetsVm != null && PresetsVm.CanConvertTo3D;
        }

        private void PresetsVm_DataChanged(ApartmentPresetData data)
        {
            ApartmentPresetDataChanged?.Invoke(data);
        }

        private void PresetsVm_StateChanged()
        {
            RelayCommand relay = ConvertTo3DCommand as RelayCommand;
            if (relay != null)
                relay.RaiseCanExecuteChanged();
        }

        private void OnUpdateApartmentMarks()
        {
            ApartmentMarksUpdateRequested?.Invoke();
        }

        private void OnConvertTo3D()
        {
            int apartmentCount = PresetsVm != null && PresetsVm.SelectedPlan != null
                ? PresetsVm.SelectedPlan.ApartmentCount
                : 0;

            if (apartmentCount > Long3DConversionApartmentWarningThreshold)
            {
                MessageBoxResult result = MessageBox.Show(
                    "Будет обработано квартир: " + apartmentCount + "\n\n" +
                    "Построение 3D для большого количества квартир может занять много времени. Продолжить?",
                    "KPLN. Менеджер квартир",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning,
                    MessageBoxResult.No);

                if (result != MessageBoxResult.Yes)
                    return;
            }

            ConvertTo3DRequested?.Invoke();
        }

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

        private static SQLiteConnection OpenConnection(string dbPath, bool readOnly)
        {
            string cs = readOnly
                ? "Data Source=" + dbPath + ";Version=3;Read Only=True;"
                : "Data Source=" + dbPath + ";Version=3;";

            var con = new SQLiteConnection(cs);
            con.Open();
            return con;
        }

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

        private void OnPick(ApartmentItemVm item)
        {
            if (item == null)
                return;

            ItemPicked?.Invoke(item.Id);
        }

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

        private void OnClose()
        {
            RequestClose?.Invoke();
        }
    }

    internal class ApartmentTypeVm
    {
        public string Name { get; set; }

        public ObservableCollection<ApartmentItemVm> Items { get; private set; }

        public ApartmentTypeVm()
        {
            Items = new ObservableCollection<ApartmentItemVm>();
        }
    }

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

    internal class RelayCommand : ICommand
    {
        private readonly Action _execute;
        private readonly Func<bool> _canExecute;

        public RelayCommand(Action e, Func<bool> canExecute = null)
        {
            _execute = e;
            _canExecute = canExecute;
        }

        public bool CanExecute(object p)
        {
            return _canExecute == null || _canExecute();
        }

        public void Execute(object p)
        {
            _execute();
        }

        public event EventHandler CanExecuteChanged;

        public void RaiseCanExecuteChanged()
        {
            EventHandler h = CanExecuteChanged;
            if (h != null)
                h(this, EventArgs.Empty);
        }
    }

    internal class RelayCommand<T> : ICommand
    {
        private readonly Action<T> _execute;
        private readonly Func<T, bool> _canExecute;

        public RelayCommand(Action<T> e, Func<T, bool> canExecute = null)
        {
            _execute = e;
            _canExecute = canExecute;
        }

        public bool CanExecute(object p)
        {
            if (_canExecute == null)
                return true;

            if (p == null)
                return false;

            return _canExecute((T)p);
        }

        public void Execute(object p)
        {
            if (p == null)
                return;

            _execute((T)p);
        }

        public event EventHandler CanExecuteChanged;

        public void RaiseCanExecuteChanged()
        {
            EventHandler h = CanExecuteChanged;
            if (h != null)
                h(this, EventArgs.Empty);
        }
    }

    internal static class ShellPreviewHelper
    {
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

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("BCC18B79-BA16-442F-80C4-8A59C30C463B")]
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
            RESIZETOFIT = 0x00,
            BIGGERSIZEOK = 0x01,
            MEMORYONLY = 0x02,
            ICONONLY = 0x04,
            THUMBNAILONLY = 0x08,
            INCACHEONLY = 0x10
        }
    }
}