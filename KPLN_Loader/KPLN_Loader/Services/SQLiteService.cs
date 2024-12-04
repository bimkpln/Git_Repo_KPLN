﻿using Autodesk.Revit.UI;
using Dapper;
using KPLN_Loader.Core.Entities;
using KPLN_Loader.Forms;
using NLog;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

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
        /// Авторизация пользователя KPLN
        /// </summary>
        /// <returns>Текущий пользователь</returns>
        internal User Authorization()
        {
            string currentDate = DateTime.Now.ToString("yyyy/MM/dd_HH:mm");
            string sysUserName = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\').Last();
            
            User currentUser = ExecuteQuery<User>($"SELECT * FROM {MainDB_Tables.Users} WHERE {nameof(User.SystemName)}='{sysUserName}';").FirstOrDefault();
            if (currentUser == null)
            {
                LoginForm loginForm = new LoginForm(SubDepartments.Where(s => s.IsAuthEnabled), false);
                if ((bool)loginForm.ShowDialog())
                {
                    int bitrixId = Task.Run(() => EnvironmentService.GetUserBitrixId_ByNameAndSurname(loginForm.CreatedWPFUser.Name, loginForm.CreatedWPFUser.Surname)).Result;
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
            ExecuteNonQuery($"UPDATE {MainDB_Tables.Users} " +
                $"SET {nameof(User.LastConnectionDate)}='{currentUser.LastConnectionDate}' WHERE {nameof(User.SystemName)}='{currentUser.SystemName}';");

            return currentUser;
        }
        #endregion

        #region Read
        /// <summary>
        /// Получить коллекцию модулей из БД
        /// </summary>
        /// <param name="currentUser">Пользователь для из БД</param>
        /// <returns>Коллекция модулей</returns>
        internal IEnumerable<Module> GetModulesForCurrentUser(User currentUser)
        {
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
        /// Получить SubDepartment пользоватлея
        /// </summary>
        /// <param name="currentUser">Пользователь для из БД</param>
        internal SubDepartment GetSubDepartmentForCurrentUser(User currentUser) => SubDepartments.FirstOrDefault(s => s.Id == currentUser.SubDepartmentId);

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
        /// Записать значение уровня одобрения от пользователя
        /// </summary>
        internal void SetLoaderDescriptionUserRank(int rate, LoaderDescription loaderDescription)
        {
            if (loaderDescription != null)
            {
                int currentRate = loaderDescription.ApprovalRate;
                int newRate = currentRate + rate;
                ExecuteNonQuery($"UPDATE {MainDB_Tables.LoaderDescriptions} " +
                    $"SET {nameof(LoaderDescription.ApprovalRate)}='{newRate}' WHERE {nameof(LoaderDescription.Id)}='{loaderDescription.Id}';");
            }
        }

        /// <summary>
        /// Обновление имени Revit-пользователя в БД
        /// </summary>
        /// <param name="userName">Текущее имя в Ревит</param>
        /// <param name="currentUser">Пользователь для проверки из БД</param>
        internal void SetRevitUserName(string userName, User currentUser)
        {
            if (currentUser.RevitUserName == null || !currentUser.RevitUserName.Equals(userName))
            {
                // Меняю объект
                currentUser.RevitUserName = userName;
                // Записываю в таблицу
                ExecuteNonQuery($"UPDATE {MainDB_Tables.Users} " +
                    $"SET {nameof(User.RevitUserName)}='{userName}' WHERE {nameof(User.SystemName)}='{currentUser.SystemName}';");
            }
        }
        #endregion

        #region Delete
        #endregion

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
