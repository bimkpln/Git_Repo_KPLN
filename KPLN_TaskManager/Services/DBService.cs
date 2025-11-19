using Dapper;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Threading;

namespace KPLN_TaskManager.Services
{
    internal class DBService
    {
        protected static void ExecuteNonQuery(string connectionString, string query, object parameters = null)
        {
            const int maxRetries = 3;
            int attempt = 0;
            int timeSleep = 100;

            while (attempt < maxRetries)
            {
                try
                {
                    using (IDbConnection connection = new SQLiteConnection(connectionString))
                    {
                        connection.Open();
                        connection.Execute(query, parameters);
                        return;
                    }
                }
                catch (SQLiteException ex) when (ex.ErrorCode == (int)SQLiteErrorCode.Busy)
                {
                    attempt++;
                    if (attempt >= maxRetries)
                    {
                        ShowDialog("[KPLN]: Ошибка работы с БД", $"База данных занята. Попытки выполнить запрос ({maxRetries} раза по {timeSleep / 1000.0} с) исчерпаны.");
                        return;
                    }

                    Thread.Sleep(timeSleep);
                }
                catch (Exception ex)
                {
                    ShowDialog("[KPLN]: Ошибка работы с БД", ex.Message);
                    return;
                }
            }
        }

        protected static int ExecuteInsertWithId(string connectionString, string query, object parameters = null)
        {
            try
            {
                using (IDbConnection connection = new SQLiteConnection(connectionString))
                {
                    connection.Open();
                    return connection.ExecuteScalar<int>(query, parameters);
                }
            }
            catch (Exception ex)
            {
                ShowDialog("[KPLN]: Ошибка работы с БД", ex.Message);
                return -1;
            }
        }

        protected static IEnumerable<T> ExecuteQuery<T>(string connectionString, string query, object parameters = null)
        {
            const int maxRetries = 3;
            int attempt = 0;
            int timeSleep = 100;

            while (attempt < maxRetries)
            {
                try
                {
                    using (IDbConnection connection = new SQLiteConnection(connectionString))
                    {
                        connection.Open();
                        return connection.Query<T>(query, parameters);
                    }
                }
                catch (SQLiteException ex) when (ex.ErrorCode == (int)SQLiteErrorCode.Busy)
                {
                    attempt++;
                    if (attempt >= maxRetries)
                    {
                        ShowDialog("[KPLN]: Ошибка работы с БД", $"База данных занята. Попытки выполнить запрос ({maxRetries} раза по {timeSleep / 1000.0} с) исчерпаны.");
                        return null;
                    }

                    Thread.Sleep(timeSleep);
                }
                catch (Exception ex)
                {
                    ShowDialog("[KPLN]: Ошибка работы с БД", ex.Message);
                    return null;
                }
            }

            return null;
        }

        /// <summary>
        /// Кастомное окно - для возможности вывода информации из другого потока (если использовать Revit API - оно просто не выведется)
        /// </summary>
        private static void ShowDialog(string title, string text)
        {
            System.Windows.Forms.Form form = new System.Windows.Forms.Form()
            {
                Text = title,
                TopMost = true,
                ShowIcon = false,
                StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
                AutoSize = true,
                AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                MinimumSize = new System.Drawing.Size(350, 150),
                MaximumSize = new System.Drawing.Size(450, 450),
            };

            System.Windows.Forms.Label textLabel = new System.Windows.Forms.Label()
            {
                Text = text,
                Font = new System.Drawing.Font("GOST Common", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204))),
                AutoSize = true,
                Dock = System.Windows.Forms.DockStyle.Fill,
                MaximumSize = new System.Drawing.Size(440, 0),
                Padding = new System.Windows.Forms.Padding(5),
            };
            form.Controls.Add(textLabel);

            System.Windows.Forms.Button confirmation = new System.Windows.Forms.Button()
            {
                Text = "Ok",
                Location = new System.Drawing.Point((form.Width - 75) / 2, 80),
                Size = new System.Drawing.Size(75, 25),
                Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) | System.Windows.Forms.AnchorStyles.Right,
            };
            confirmation.Click += (sender, e) => { form.Close(); };
            form.Controls.Add(confirmation);

            form.ShowDialog();
        }
    }
}
