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
    internal class SQLiteService
    {
        private readonly Logger _logger;
        private readonly string _dbPath;

        public SQLiteService(Logger logger, string dbPath)
        {
            _logger = logger;
            _dbPath = "Data Source=" + dbPath + "; Version=3;";
        }

        /// <summary>
        /// Текущий пользователь из БД
        /// </summary>
        public User CurrentUser { get; private set; }

        /// <summary>
        /// Коллекция отделов из БД
        /// </summary>
        public IEnumerable<SubDepartment> SubDepartments { get; private set; }

        /// <summary>
        /// Авторизация пользователя KPLN
        /// </summary>
        public void Authorization()
        {
            string currentDate = DateTime.Now.ToString("yyyy/MM/dd_HH:mm");
            string sysUserName = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\').Last();
            CurrentUser = ExecuteQuery<User>($"SELECT * FROM Users WHERE {nameof(User.SystemName)}='{sysUserName}';").FirstOrDefault();
            SubDepartments = ExecuteQuery<SubDepartment>("SELECT * FROM SubDepartments;");

            if (CurrentUser == null)
            {
                LoginForm loginForm = new LoginForm(SubDepartments.Where(s => s.IsAuthEnabled.ToLower().Equals("true")));
                if ((bool)loginForm.ShowDialog())
                {
                    User user = new User()
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
                        user);

                    CurrentUser = user;
                    _logger.Info($"Пользователь {CurrentUser.SystemName} успешно создан и записан в БД!");
                }
            }
            else
                _logger.Info($"Пользователь {CurrentUser.SystemName} успешно определен!");
            
            CurrentUser.LastConnectionDate = currentDate;
            ExecuteNonQuery($"UPDATE Users SET {nameof(User.LastConnectionDate)}='{CurrentUser.LastConnectionDate}' WHERE {nameof(User.SystemName)}='{CurrentUser.SystemName}';");
        }

        /// <summary>
        /// Получить коллекцию модулей из БД
        /// </summary>
        public IEnumerable<Module> GetModulesForCurrentUser()
        {
            IEnumerable<Module> modules;
            if (CurrentUser.IsDebugMode.ToLower().Equals("true"))
            {
                modules = ExecuteQuery<Module>($"SELECT * FROM Modules " +
                    $"WHERE {nameof(Module.IsDebugMode)}='True';");
            }
            else
            {
                modules = ExecuteQuery<Module>($"SELECT * FROM Modules " +
                        $"WHERE {nameof(Module.IsEnabled)}='True' " +
                        $"AND ({nameof(Module.SubDepartmentId)}=1 OR {nameof(Module.SubDepartmentId)}={CurrentUser.SubDepartmentId});");
            }
            return modules;
        }

        /// <summary>
        /// Получить описание окна загрузки из БД
        /// </summary>
        /// <returns></returns>
        public string GetDescriptionForCurrentUser() 
        {
            IEnumerable<LoaderDescription> loaderDescriptions;
            SubDepartment bimDep = SubDepartments.Where(s => s.Code.ToLower().Contains("bim")).FirstOrDefault();
            if (CurrentUser.SubDepartmentId == bimDep.Id)
            {
                loaderDescriptions = ExecuteQuery<LoaderDescription>($"SELECT * FROM LoaderDescriptions;");
            }
            else
            {
                loaderDescriptions = ExecuteQuery<LoaderDescription>($"SELECT * FROM LoaderDescriptions " +
                        $"WHERE ({nameof(LoaderDescription.SubDepartmentId)}=1 OR {nameof(Module.SubDepartmentId)}={CurrentUser.SubDepartmentId});");
            }
            
            Random rand = new Random();
            int index = rand.Next(loaderDescriptions.Count() + 1);
            return loaderDescriptions.ElementAt(index).Description;

        } 

        /// <summary>
        /// Обновление имени Revit-пользователя в БД
        /// </summary>
        /// <param name="userName"></param>
        public void SetRevitUserName(string userName)
        {
            if (!CurrentUser.RevitUserName.Equals(userName))
                ExecuteNonQuery($"UPDATE Users SET {nameof(User.RevitUserName)}='{userName}' WHERE {nameof(User.SystemName)}='{CurrentUser.SystemName}';");
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
