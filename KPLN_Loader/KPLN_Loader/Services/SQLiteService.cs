using Autodesk.Revit.UI;
using Dapper;
using KPLN_Loader.Core.Entities;
using KPLN_Loader.Forms;
using NLog;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace KPLN_Loader.Services
{
    /// <summary>
    /// Сервис для работы с БД
    /// </summary>
    internal sealed class SQLiteService
    {
        private readonly Logger _logger;
        private readonly string _dbPath;
        private IEnumerable<SubDepartment> _subDepartments;

        internal SQLiteService(Logger logger, string dbPath)
        {
            _logger = logger;
            _dbPath = "Data Source=" + dbPath + "; Version=3;";
        }

        /// <summary>
        /// Кэширование коллекции отделов из БД
        /// </summary>
        private IEnumerable<SubDepartment> SubDepartments
        {
            get
            {
                if (_subDepartments == null)
                {
                    _subDepartments = ExecuteQuery<SubDepartment>($"SELECT * FROM {MainDB_Tables.SubDepartments};");
                }
                return _subDepartments;
            }
        }

        #region Create
        /// <summary>
        /// БД: Авторизация пользователя KPLN
        /// </summary>
        /// <returns>Текущий пользователь</returns>
        internal User Authorization(EnvironmentService envService)
        {
            _logger.Info("БД: Авторизация пользователя KPLN");

            string currentDate = DateTime.Now.ToString("yyyy/MM/dd_HH:mm");
            string sysUserName = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\').Last();

            User currentUser = ExecuteQuery<User>($"SELECT * FROM {MainDB_Tables.Users} WHERE {nameof(User.SystemName)}='{sysUserName}';").FirstOrDefault();
            if (currentUser == null)
            {
                LoginForm loginForm = new LoginForm(SubDepartments.Where(s => s.IsAuthEnabled), false);
                if ((bool)loginForm.ShowDialog())
                {
                    int bitrixId = Task.Run(() => envService.GetUserBitrixId_ByNameAndSurname(loginForm.CreatedWPFUser.Name, loginForm.CreatedWPFUser.Surname)).Result;
                    currentUser = new User()
                    {
                        SystemName = sysUserName,
                        Name = loginForm.CreatedWPFUser.Name,
                        Surname = loginForm.CreatedWPFUser.Surname,
                        Company = loginForm.CreatedWPFUser.Company,
                        SubDepartmentId = loginForm.CreatedWPFUser.SubDepartment.Id,
                        RegistrationDate = currentDate,
                        IsUserRestricted = true,
                        BitrixUserID = bitrixId,
                    };

                    ExecuteNonQuery($"INSERT INTO {MainDB_Tables.Users} " +
                            $"({nameof(User.SystemName)}, {nameof(User.Name)}, {nameof(User.Surname)}, {nameof(User.SubDepartmentId)}, {nameof(User.RegistrationDate)}, {nameof(User.BitrixUserID)}) " +
                            $"VALUES (@{nameof(User.SystemName)}, @{nameof(User.Name)}, @{nameof(User.Surname)}, @{nameof(User.SubDepartmentId)}, @{nameof(User.RegistrationDate)}, @{nameof(User.BitrixUserID)});",
                        currentUser);

                    _logger.Info($"Пользователь {currentUser.SystemName}: " +
                        $"{currentUser.Surname} {currentUser.Name} из отдела {GetSubDepartmentForCurrentUser(currentUser).Code} " +
                        $"успешно создан и записан в БД!");
                }
                else
                {
                    TaskDialog.Show(
                        "Ошибка",
                        "Пользователь отменил ввод данных. Загрузка завершена с ошибкой. Перезапустите Revit и заполните все строки в окне регистрации",
                        TaskDialogCommonButtons.Cancel);

                    return null;
                }
            }
            else
            {
                _logger.Info($"Пользователь {currentUser.SystemName}: " +
                    $"{currentUser.Surname} {currentUser.Name} из отдела {GetSubDepartmentForCurrentUser(currentUser).Code} " +
                    $"успешно определен!");

            }

            if (currentUser.IsDebugMode)
                _logger.Info("ВЫБРАН СТАТУС ЗАПУСКА - DEBUG");

            currentUser.LastConnectionDate = currentDate;

            return currentUser;
        }
        #endregion

        #region Read
        /// <summary>
        /// БД: Получить коллекцию модулей 
        /// </summary>
        /// <param name="currentUser">Пользователь для из БД</param>
        /// <returns>Коллекция модулей</returns>
        internal IEnumerable<Module> GetModulesForCurrentUser(User currentUser)
        {
            _logger.Info("БД: Получить коллекцию модулей");
            
            IEnumerable<Module> modules;
            // Модули-библиотеки статуса Debug хранятся в спец. папках "Debug", а далее - по аналогии с остальными модулями
            if (currentUser.IsDebugMode)
            {
                modules = ExecuteQuery<Module>($"SELECT * FROM {MainDB_Tables.Modules} " +
                    $"WHERE {nameof(Module.IsEnabled)}='True'" +
                    $"AND ({nameof(Module.IsLibraryModule)}='True' OR {nameof(Module.IsDebugMode)}='True');");
            }
            else
            {
                modules = ExecuteQuery<Module>($"SELECT * FROM {MainDB_Tables.Modules} " +
                    $"WHERE {nameof(Module.IsEnabled)}='True' " +
                    $"AND ({nameof(Module.SubDepartmentId)}=1 OR {nameof(Module.SubDepartmentId)}={currentUser.SubDepartmentId}) " +
                    $"AND ({nameof(Module.IsLibraryModule)}='True' OR {nameof(Module.IsDebugMode)}='False');");
            }

            return modules;
        }

        /// <summary>
        /// БД: Получить SubDepartment пользователя
        /// </summary>
        /// <param name="currentUser">Пользователь для из БД</param>
        internal SubDepartment GetSubDepartmentForCurrentUser(User currentUser)
        {
            _logger.Info("БД: Получить SubDepartment пользователя");

            return SubDepartments.FirstOrDefault(s => s.Id == currentUser.SubDepartmentId);
        }

        /// <summary>
        /// БД: Получить описание окна загрузки
        /// </summary>
        /// <param name="currentUser">Пользователь для из БД</param>
        /// <returns>Строка из БД</returns>
        internal LoaderDescription GetDescriptionForCurrentUser(User currentUser)
        {
            _logger.Info("БД: Получить описание окна загрузки");

            IEnumerable<LoaderDescription> loaderDescriptions;
            SubDepartment subDepartment = GetSubDepartmentForCurrentUser(currentUser);
            if (subDepartment.Code.ToLower().Contains("bim"))
            {
                loaderDescriptions = ExecuteQuery<LoaderDescription>($"SELECT * FROM {MainDB_Tables.LoaderDescriptions};");
            }
            else
            {
                loaderDescriptions = ExecuteQuery<LoaderDescription>($"SELECT * FROM {MainDB_Tables.LoaderDescriptions} " +
                        $"WHERE ({nameof(LoaderDescription.SubDepartmentId)}=1 OR {nameof(Module.SubDepartmentId)}={currentUser.SubDepartmentId});");
            }

            Random rand = new Random();
            int index = rand.Next(loaderDescriptions.Count() - 1);
            return loaderDescriptions.ElementAt(index);
        }
        #endregion

        #region Update
        /// <summary>
        /// БД: Обновить данные по последнему запуску
        /// </summary>
        internal bool SetUserLastConnectionDate(User currentUser)
        {
            _logger.Info("БД: Обновить данные по последнему запуску");

            try
            {
                ExecuteNonQuery($"UPDATE {MainDB_Tables.Users} " +
                    $"SET {nameof(User.LastConnectionDate)}='{currentUser.LastConnectionDate}' WHERE {nameof(User.SystemName)}='{currentUser.SystemName}';");
                
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// БД: Записать значение уровня одобрения от пользователя
        /// </summary>
        internal void SetLoaderDescriptionUserRank(MainDB_LoaderDescriptions_RateType rateType, LoaderDescription loaderDescription)
        {
            _logger.Info("БД: Записать значение уровня одобрения от пользователя");

            if (loaderDescription != null)
            {
                switch (rateType)
                {
                    case MainDB_LoaderDescriptions_RateType.Approval:
                        int apprRate = loaderDescription.ApprovalRate;
                        ExecuteNonQuery($"UPDATE {MainDB_Tables.LoaderDescriptions} " +
                            $"SET {nameof(LoaderDescription.ApprovalRate)}='{++apprRate}' WHERE {nameof(LoaderDescription.Id)}='{loaderDescription.Id}';");
                        break;
                    case MainDB_LoaderDescriptions_RateType.Disapproval:
                        int disapprRate = loaderDescription.DisapprovalRate;
                        ExecuteNonQuery($"UPDATE {MainDB_Tables.LoaderDescriptions} " +
                            $"SET {nameof(LoaderDescription.DisapprovalRate)}='{--disapprRate}' WHERE {nameof(LoaderDescription.Id)}='{loaderDescription.Id}';");
                        break;
                }
            }
        }

        /// <summary>
        /// БД: Обновление имени Revit-пользователя
        /// </summary>
        /// <param name="userName">Текущее имя в Ревит</param>
        /// <param name="currentUser">Пользователь для проверки из БД</param>
        internal void SetRevitUserName(string userName, User currentUser)
        {
            if (currentUser.RevitUserName == null || !currentUser.RevitUserName.Equals(userName))
            {
                try
                {
                    _logger.Info($"БД: Обновление имени Revit-пользователя с {currentUser.RevitUserName} на {userName}");
                    
                    // Меняю объект
                    currentUser.RevitUserName = userName;
                
                    // Записываю в таблицу
                    ExecuteNonQuery($"UPDATE {MainDB_Tables.Users} " +
                        $"SET {nameof(User.RevitUserName)}='{userName}' WHERE {nameof(User.SystemName)}='{currentUser.SystemName}';");
                }
                catch (Exception ex)
                {
                    _logger.Error($"БД: Не удалось обновить имени Revit-пользователя. Ошибка: {ex.Message}");
                }
            }
        }
        #endregion

        #region Delete
        #endregion

        private void ExecuteNonQuery(string query, object parameters = null)
        {
            const int maxRetries = 3;
            int attempt = 0;
            int timeSleep = 1000;

            while (attempt < maxRetries)
            {
                try
                {
                    using (IDbConnection connection = new SQLiteConnection(_dbPath))
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
                        string errorMsg = $"База данных занята. Попытки выполнить запрос ({maxRetries} раза по {timeSleep / 1000} с) исчерпаны.";

                        ShowDialog("[KPLN]: Ошибка работы с БД", errorMsg);

                        throw new Exception(errorMsg);
                    }

                    Thread.Sleep(timeSleep);
                }
            }
        }

        private IEnumerable<T> ExecuteQuery<T>(string query, object parameters = null)
        {
            const int maxRetries = 3;
            int attempt = 0;
            int timeSleep = 1000;

            while (attempt < maxRetries)
            {
                try
                {
                    using (IDbConnection connection = new SQLiteConnection(_dbPath))
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
                        string errorMsg = $"База данных занята. Попытки выполнить запрос ({maxRetries} раза по {timeSleep / 1000} с) исчерпаны.";

                        ShowDialog("[KPLN]: Ошибка работы с БД", errorMsg);

                        throw new Exception(errorMsg);
                    }

                    Thread.Sleep(timeSleep);
                }
            }

            throw new Exception("Не удалось получить результатов. Отправь разработчику");
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
