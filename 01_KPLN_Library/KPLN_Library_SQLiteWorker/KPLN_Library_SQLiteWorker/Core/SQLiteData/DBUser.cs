using KPLN_Library_SQLiteWorker.Core.SQLiteData.Abstractions;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace KPLN_Library_SQLiteWorker.Core.SQLiteData
{
    /// <summary>
    /// Класс пользователя KPLN
    /// </summary>
    public class DBUser : IDBEntity
    {
        #region Столбцы из БД
        [Key]
        public int Id { get; set; }

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

        /// <summary>
        /// Ограничить работу пользователя вкл/выкл (True/False). В БД тип данных текст, преобразование происходит в Dapper
        /// </summary>
        public bool IsUserRestricted { get; set; }

        /// <summary>
        /// ID пользователя Bitrix
        /// </summary>
        public int BitrixUserID { get; set; }
        #endregion

        /// <summary>
        /// Привязка к БД из DB_Enumerator
        /// </summary>
        public static DB_Enumerator CurrentDB { get; } = DB_Enumerator.Users;
    }
}
