using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data.SQLite;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Win32;

namespace KPLN_Tools.Forms
{
    public partial class ApartmentManagerWindow : Window
    {
        public int SelectedApartmentId { get; private set; }
        public int _nDep;

        public ApartmentManagerWindow(int nDep)
        {
            InitializeComponent();
            _nDep = nDep;

            ApartmentManagerVm vm = new ApartmentManagerVm(_nDep);
            vm.ItemPicked += Vm_ItemPicked;
            vm.RequestClose += Vm_RequestClose;

            DataContext = vm;
        }

        private void Vm_ItemPicked(int id)
        {
            SelectedApartmentId = id;
            DialogResult = true;
            Close();
        }

        private void Vm_RequestClose()
        {
            DialogResult = false;
            Close();
        }
    }

    internal class ApartmentManagerVm : INotifyPropertyChanged
    {
        private const string DbPath = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_ApartmentManager.db";

        public event Action RequestClose;
        public event Action<int> ItemPicked;

        public bool IsDep8 { get; private set; }

        public ObservableCollection<ApartmentTypeVm> ApartmentTypes { get; private set; }
        private ApartmentTypeVm _selectedType;
        public ApartmentTypeVm SelectedType
        {
            get { return _selectedType; }
            set
            {
                if (!object.ReferenceEquals(_selectedType, value))
                {
                    _selectedType = value;
                    OnPropertyChanged();
                    LoadItems();
                }
            }
        }

        public ICommand PickItemCommand { get; private set; }
        public ICommand CloseCommand { get; private set; }
        public ICommand RefreshDbCommand { get; private set; }
        public ICommand UploadImageCommand { get; private set; }

        public ApartmentManagerVm(int nDep)
        {
            IsDep8 = (nDep == 8);

            ApartmentTypes = new ObservableCollection<ApartmentTypeVm>();

            PickItemCommand = new RelayCommand<ApartmentItemVm>(OnPick);
            CloseCommand = new RelayCommand(OnClose);

            RefreshDbCommand = new RelayCommand(OnRefreshDb);
            UploadImageCommand = new RelayCommand<ApartmentItemVm>(OnUploadImage);

            LoadTypes();
        }













        // ОБНОВИТЬ БД
        private void OnRefreshDb()
        {           
            LoadTypes();
            if (SelectedType != null)
                LoadItems();
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
                using (var con = OpenConnection(DbPath, readOnly: true))
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

                            if (string.IsNullOrWhiteSpace(atype))
                                continue;

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
                using (var con = OpenConnection(DbPath, readOnly: true))
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
            if (item == null)
                return;

            if (!IsDep8)
                return;

            var ofd = new OpenFileDialog();
            ofd.Title = "Выберите изображение планировки";
            ofd.Filter = "Изображения (*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg|Все файлы (*.*)|*.*";
            ofd.Multiselect = false;

            bool? ok = ofd.ShowDialog();
            if (ok != true)
                return;

            string filePath = ofd.FileName;
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                return;

            try
            {
                byte[] bytes = File.ReadAllBytes(filePath);

                using (var con = OpenConnection(DbPath, readOnly: false))
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

        private static SQLiteConnection OpenConnection(string dbPath, bool readOnly)
        {
            string cs = readOnly
                ? ("Data Source=" + dbPath + ";Version=3;Read Only=True;")
                : ("Data Source=" + dbPath + ";Version=3;");

            var con = new SQLiteConnection(cs);
            con.Open();
            return con;
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

        private void OnPick(ApartmentItemVm item)
        {
            if (item == null)
                return;

            ItemPicked?.Invoke(item.Id);
        }

        private void OnClose()
        {
            RequestClose?.Invoke();
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string p = null)
        {
            var h = PropertyChanged;
            if (h != null) h(this, new PropertyChangedEventArgs(p));
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
            if (h != null) h(this, new PropertyChangedEventArgs(p));
        }
    }

    internal class RelayCommand : ICommand
    {
        private readonly Action _execute;

        public RelayCommand(Action e)
        {
            _execute = e;
        }

        public bool CanExecute(object p) { return true; }
        public void Execute(object p) { _execute(); }
        public event EventHandler CanExecuteChanged;
    }

    internal class RelayCommand<T> : ICommand
    {
        private readonly Action<T> _execute;

        public RelayCommand(Action<T> e)
        {
            _execute = e;
        }

        public bool CanExecute(object p) { return true; }

        public void Execute(object p)
        {
            if (p == null) return;
            _execute((T)p);
        }

        public event EventHandler CanExecuteChanged;
    }

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
}