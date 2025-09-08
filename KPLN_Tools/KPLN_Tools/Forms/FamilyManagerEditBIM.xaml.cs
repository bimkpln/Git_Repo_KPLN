using System;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;

namespace KPLN_Tools.Forms
{
    public partial class FamilyManagerEditBIM : Window
    {
        private const string DB_PATH = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_FamilyManager.db";

        private string _originalStatus;

        // Класс БД
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

        // Основной конструктор
        public FamilyManagerEditBIM(string idText)
        {
            InitializeComponent();

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
            SetupStatusCombo(_originalStatus);
        }

        // БД. Чтение
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

















        // XAML. Отмена
        private void BtnCancel_OnClick(object sender, RoutedEventArgs e)
        {
            Close();
        }


        private void ButtonSaveDB_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
