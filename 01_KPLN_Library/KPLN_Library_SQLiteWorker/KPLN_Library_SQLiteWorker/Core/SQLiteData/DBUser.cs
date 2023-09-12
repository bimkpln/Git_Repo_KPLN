﻿using KPLN_Library_SQLiteWorker.Core.SQLiteData.Abstractions;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace KPLN_Library_SQLiteWorker.Core.SQLiteData
{
    /// <summary>
    /// Класс пользователя KPLN
    /// </summary>
    public class DBUser : IDBEntity
    {
        [Key]
        public int Id { get; set; }

        public DB_Enumerator CurrentDB { get; set; }

        /// <summary>
        /// Системное имя пользователя (Windows)
        /// </summary>
        public string SystemName { get; set; }

        /// <summary>
        /// Имя пользователя
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Фамилия пользователя
        /// </summary>
        public string Surname { get; set; }

        /// <summary>
        /// Отдел пользователя
        /// </summary>
        [ForeignKey(nameof(DBSubDepartment))]
        public int SubDepartmentId { get; set; }

        /// <summary>
        /// Дата регистрации
        /// </summary>
        public string RegistrationDate { get; set; }

        /// <summary>
        /// Дата подключения
        /// </summary>
        public string LastConnectionDate { get; set; }

        /// <summary>
        /// Имя Revit-пользователя 
        /// </summary>
        public string RevitUserName { get; set; }

        /// <summary>
        /// Режим отладки вкл/выкл (True/False). В БД тип данных текст, преобразование происходит в Dapper
        /// </summary>
        public bool IsDebugMode { get; set; }
    }
}
