using System;
using System.Collections.Generic;
using System.Data.SQLite; 
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Text.RegularExpressions;

namespace KPLN_Tools.Forms
{
    public partial class FamilyManagerSettingDB : Window
    {
        private const string DB_PATH = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_FamilyManager.db";

        public FamilyManagerSettingDB()
        {
            InitializeComponent();
            Loaded += FamilyManagerSettingDB_Loaded;
        }

        private void FamilyManagerSettingDB_Loaded(object sender, RoutedEventArgs e)
        {
            string currentUser = Environment.UserName;

            if (!string.Equals(currentUser, "rtuleninov", StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show(
                    $"Текущий пользователь: {currentUser}\nДоступ разрешён только для rtuleninov.",
                    "Отказано в доступе",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                DialogResult = false; 
                Close();       
            }
        }

        private void OnSelectionChanged(object sender, RoutedEventArgs e)
        {
            btnDelete.IsEnabled = cbInfo.IsChecked == true || cbImage.IsChecked == true;
        }

        private async void Delete_Click(object sender, RoutedEventArgs e)
        {
            bool delInfo = cbInfo.IsChecked == true;
            bool delImage = cbImage.IsChecked == true;

            if (!delInfo && !delImage)
                return;

            var confirm = MessageBox.Show(BuildConfirmText(delInfo, delImage), "Подтверждение удаления", MessageBoxButton.OKCancel, MessageBoxImage.Warning);

            if (confirm != MessageBoxResult.OK)
                return;

            btnDelete.IsEnabled = false;

            try
            {
                await DoDeletionAsync(delInfo, delImage);

                MessageBox.Show("Удаление выполнено.", "Готово",
                                MessageBoxButton.OK, MessageBoxImage.Information);

                DialogResult = true; 
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при удалении:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                btnDelete.IsEnabled = true;
            }
        }

        private static string BuildConfirmText(bool delInfo, bool delImage)
        {
            string list = "";
            if (delInfo) list += "• Информация\n";
            if (delImage) list += "• Изображение\n";
            return "Удалить следующие элементы (для всех записей)?\n" + list.TrimEnd();
        }

        private Task DoDeletionAsync(bool delInfo, bool delImage)
        {
            return Task.Run(() =>
            {
                if (!File.Exists(DB_PATH))
                    throw new FileNotFoundException("Не найден файл базы данных.", DB_PATH);

                var setParts = new List<string>();
                if (delInfo) setParts.Add("IMPORT_INFO = NULL");
                if (delImage) setParts.Add("IMAGE = NULL");

                if (setParts.Count == 0)
                    return;

                string sql = $"UPDATE FamilyManager SET {string.Join(", ", setParts)};";

                var csb = new SQLiteConnectionStringBuilder
                {
                    DataSource = DB_PATH,
                    Version = 3,
                    FailIfMissing = true
                };

                using (var conn = new SQLiteConnection(csb.ToString()))
                {
                    conn.Open();
                    using (var tran = conn.BeginTransaction())
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = tran;
                        cmd.CommandText = sql;
                        int affected = cmd.ExecuteNonQuery();
                        tran.Commit();
                    }
                }
            });
        }
        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }      
    }
}
