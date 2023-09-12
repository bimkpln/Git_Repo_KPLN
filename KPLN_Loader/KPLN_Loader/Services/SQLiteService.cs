using Dapper;
using KPLN_Loader.Core.SQLiteData;
using KPLN_Loader.Forms;
using NLog;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Linq;

namespace KPLN_Loader.Services
{
    /// <summary>
    /// Сервис для работы с БД
    /// </summary>
    internal sealed class SQLiteService
    {
        private IEnumerable<SubDepartment> _subDepartments;
        private readonly Logger _logger;
        private readonly string _dbPath;

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
                    _subDepartments = ExecuteQuery<SubDepartment>("SELECT * FROM SubDepartments;");
                }
                return _subDepartments;
            }
        }

        /// <summary>
        /// Авторизация пользователя KPLN
        /// </summary>
        /// <returns>Текущий пользователь</returns>
        internal User Authorization()
        {
            string currentDate = DateTime.Now.ToString("yyyy/MM/dd_HH:mm");
            string sysUserName = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\').Last();
            User currentUser = ExecuteQuery<User>($"SELECT * FROM Users WHERE {nameof(User.SystemName)}='{sysUserName}';").FirstOrDefault();

            if (currentUser == null)
            {
                LoginForm loginForm = new LoginForm(SubDepartments.Where(s => s.IsAuthEnabled));
                if ((bool)loginForm.ShowDialog())
                {
                    currentUser = new User()
                    {
                        SystemName = sysUserName,
                        Name = loginForm.UserName,
                        Surname = loginForm.Surname,
                        SubDepartmentId = loginForm.CurrentSubDepartment.Id,
                        RegistrationDate = currentDate,
                    };
                    
                    ExecuteNonQuery($"INSERT INTO Users " +
                            $"({nameof(User.SystemName)}, {nameof(User.Name)}, {nameof(User.Surname)}, {nameof(User.SubDepartmentId)}, {nameof(User.RegistrationDate)}) " +
                            $"VALUES (@{nameof(User.SystemName)}, @{nameof(User.Name)}, @{nameof(User.Surname)}, @{nameof(User.SubDepartmentId)}, @{nameof(User.RegistrationDate)});",
                        currentUser);

                    _logger.Info($"Пользователь {currentUser.SystemName}: " +
                        $"{currentUser.Surname} {currentUser.Name} из отдела {GetSubDepartmentForCurrentUser(currentUser).Code} " +
                        $"успешно создан и записан в БД!");
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
            ExecuteNonQuery($"UPDATE Users SET {nameof(User.LastConnectionDate)}='{currentUser.LastConnectionDate}' WHERE {nameof(User.SystemName)}='{currentUser.SystemName}';");

            return currentUser;
        }

        /// <summary>
        /// Получить коллекцию модулей из БД
        /// </summary>
        /// <param name="currentUser">Пользователь для из БД</param>
        /// <returns>Коллекция модулей</returns>
        internal IEnumerable<Module> GetModulesForCurrentUser(User currentUser)
        {
            IEnumerable<Module> modules;
            // Модули-библиотеки нужны при любом режиме, для них статус IsDebugMode - не играет роли
            if (currentUser.IsDebugMode)
            {
                modules = ExecuteQuery<Module>("SELECT * FROM Modules " +
                    $"WHERE {nameof(Module.IsEnabled)}='True'" +
                    $"AND ({nameof(Module.IsLibraryModule)}='True' OR {nameof(Module.IsDebugMode)}='True');");
            }
            else
            {
                modules = ExecuteQuery<Module>("SELECT * FROM Modules " +
                        $"WHERE {nameof(Module.IsEnabled)}='True' " +
                        $"AND ({nameof(Module.SubDepartmentId)}=1 OR {nameof(Module.SubDepartmentId)}={currentUser.SubDepartmentId}) " +
                        $"AND ({nameof(Module.IsLibraryModule)}='True' OR {nameof(Module.IsDebugMode)}='False');");
            }
            return modules;
        }

        /// <summary>
        /// Получить SubDepartment пользоватлея
        /// </summary>
        /// <param name="currentUser">Пользователь для из БД</param>
        internal SubDepartment GetSubDepartmentForCurrentUser(User currentUser) => SubDepartments.Where(s => s.Id == currentUser.SubDepartmentId).FirstOrDefault();

        /// <summary>
        /// Получить описание окна загрузки из БД
        /// </summary>
        /// <param name="currentUser">Пользователь для из БД</param>
        /// <returns>Строка из БД</returns>
        internal LoaderDescription GetDescriptionForCurrentUser(User currentUser) 
        {
            IEnumerable<LoaderDescription> loaderDescriptions;
            SubDepartment subDepartment = GetSubDepartmentForCurrentUser(currentUser);
            if (subDepartment.Code.ToLower().Contains("bim"))
            {
                loaderDescriptions = ExecuteQuery<LoaderDescription>($"SELECT * FROM LoaderDescriptions;");
            }
            else
            {
                loaderDescriptions = ExecuteQuery<LoaderDescription>($"SELECT * FROM LoaderDescriptions " +
                        $"WHERE ({nameof(LoaderDescription.SubDepartmentId)}=1 OR {nameof(Module.SubDepartmentId)}={currentUser.SubDepartmentId});");
            }
            
            Random rand = new Random();
            int index = rand.Next(loaderDescriptions.Count() - 1);
            return loaderDescriptions.ElementAt(index);
        }

        /// <summary>
        /// Обновление имени Revit-пользователя в БД
        /// </summary>
        /// <param name="userName">Текущее имя в Ревит</param>
        /// <param name="currentUser">Пользователь для проверки из БД</param>
        internal void SetRevitUserName(string userName, User currentUser)
        {
            if (currentUser.RevitUserName == null || !currentUser.RevitUserName.Equals(userName))
                ExecuteNonQuery($"UPDATE Users SET {nameof(User.RevitUserName)}='{userName}' WHERE {nameof(User.SystemName)}='{currentUser.SystemName}';");
        }

        private void ExecuteNonQuery(string query, object parameters = null)
        {
            using (IDbConnection connection = new SQLiteConnection(_dbPath))
            {
                connection.Open();
                connection.Execute(query, parameters);
            }
        }

        private IEnumerable<T> ExecuteQuery<T>(string query, object parameters = null)
        {
            using (IDbConnection connection = new SQLiteConnection(_dbPath))
            {
                connection.Open();
                return connection.Query<T>(query, parameters);
            }
        }
    }
}
