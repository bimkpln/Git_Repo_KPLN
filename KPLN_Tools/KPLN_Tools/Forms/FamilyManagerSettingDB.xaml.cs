using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Data.SQLite; 

namespace KPLN_Tools.Forms
{
    public partial class FamilyManagerSettingDB : Window
    {
        private const string DB_PATH = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_FamilyManager.db";

        public FamilyManagerSettingDB()
        {
            InitializeComponent();
        }

        private void OnSelectionChanged(object sender, RoutedEventArgs e)
        {
            btnDelete.IsEnabled = cbDept.IsChecked == true || cbInfo.IsChecked == true || cbImage.IsChecked == true;
        }

        private async void Delete_Click(object sender, RoutedEventArgs e)
        {
            bool delDept = cbDept.IsChecked == true;
            bool delInfo = cbInfo.IsChecked == true;
            bool delImage = cbImage.IsChecked == true;

            if (!delDept && !delInfo && !delImage)
                return;

            var confirm = MessageBox.Show(BuildConfirmText(delDept, delInfo, delImage), "Подтверждение удаления", MessageBoxButton.OKCancel, MessageBoxImage.Warning);

            if (confirm != MessageBoxResult.OK)
                return;

            btnDelete.IsEnabled = false;

            try
            {
                await DoDeletionAsync(delDept, delInfo, delImage);

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

        private static string BuildConfirmText(bool delDept, bool delInfo, bool delImage)
        {
            string list = "";
            if (delDept) list += "• Отдел\n";
            if (delInfo) list += "• Информация\n";
            if (delImage) list += "• Изображение\n";
            return "Удалить следующие элементы (для всех записей)?\n" + list.TrimEnd();
        }

        private Task DoDeletionAsync(bool delDept, bool delInfo, bool delImage)
        {
            return Task.Run(() =>
            {
                if (!File.Exists(DB_PATH))
                    throw new FileNotFoundException("Не найден файл базы данных.", DB_PATH);

                var setParts = new List<string>();
                if (delDept) setParts.Add("DEPARTAMENT = NULL");
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
