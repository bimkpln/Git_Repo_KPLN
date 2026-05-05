using Dapper;
using KPLN_Library_DBWorker.Core;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Threading;

namespace KPLN_Library_DBWorker.FactoryParts.Common
{
    /// <summary>
    /// Абстрактный класс для сервисов работы с БД, чтобы не повторять код обработки ошибок и повторов запросов
    /// </summary>
    public abstract class DbServiceAbstr<TConnection, TDbException>
        where TConnection : IDbConnection
        where TDbException : Exception
    {
        private protected readonly string _connectionString;
        private protected string _dbTableName;

        /// <summary>
        /// Конструктор только для наследников
        /// </summary>
        /// <param name="connectionString">Строка подключения</param>
        private protected DbServiceAbstr(string connectionString, DBEnumerator dbEnumerator)
        {
            _connectionString = connectionString;
            _dbTableName = dbEnumerator.ToString();
        }

        /// <summary>
        /// Создать инстанс подключения
        /// </summary>
        protected abstract TConnection CreateConnection(string connectionString);

        /// <summary>
        /// Определить, можно ли повторить запрос
        /// </summary>
        protected abstract bool IsRetryable(TDbException ex);

        /// <summary>
        /// Человекочитаемое имя провайдера БД
        /// </summary>
        protected abstract string DbProviderName { get; }

        /// <summary>
        /// SQL-литерал логического TRUE для текущего провайдера
        /// </summary>
        protected virtual string SqlTrueLiteral => "'True'";

        /// <summary>
        /// SQL-литерал логического FALSE для текущего провайдера
        /// </summary>
        protected virtual string SqlFalseLiteral => "'False'";

        private protected void ExecuteNonQuery(string query, object parameters = null)
        {
            const int maxRetries = 3;
            int attempt = 0;
            int timeSleep = 2000;

            while (attempt < maxRetries)
            {
                try
                {
                    using (IDbConnection connection = CreateConnection(_connectionString))
                    {
                        connection.Open();
                        using (IDbTransaction trans = connection.BeginTransaction())
                        {
                            connection.Execute(query, parameters, trans);
                            trans.Commit();
                        }
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

        private protected IEnumerable<T> ExecuteQuery<T>(string query, object parameters = null)
        {
            const int maxRetries = 3;
            int attempt = 0;
            int timeSleep = 1000;

            while (attempt < maxRetries)
            {
                try
                {
                    using (IDbConnection connection = CreateConnection(_connectionString))
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
        /// Обработка запроса на несколько элементов сразу
        /// ВАЖНО1: Таблица в БД ОБЯЗАНА иметь UNIQUE INDEX, иначе нет точки выхода, и данные будут писаться БЕЗКОНЕЧНО
        /// ВАЖНО2: SQLite запрос должен содержать блок игнорирования, иначе упадёт ошибкой
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="query"></param>
        /// <param name="items"></param>
        private protected void ExecuteBulkInsert<T>(string query, IEnumerable<T> items)
        {
            const int maxRetries = 3;
            int attempt = 0;
            int timeSleep = 1000;

            while (attempt < maxRetries)
            {
                try
                {
                    using (IDbConnection connection = CreateConnection(_connectionString))
                    {
                        connection.Open();
                        // Открываю транзакцию, чтобы не записывать в БД строка за строкой
                        using (IDbTransaction trans = connection.BeginTransaction())
                        {
                            connection.Execute(query, items, trans);
                            trans.Commit();
                        }
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

            return;
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
