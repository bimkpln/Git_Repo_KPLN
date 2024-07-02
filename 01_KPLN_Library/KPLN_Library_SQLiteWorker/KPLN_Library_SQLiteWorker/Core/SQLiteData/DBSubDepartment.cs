using KPLN_Library_SQLiteWorker.Core.SQLiteData.Abstractions;
using System.ComponentModel.DataAnnotations;

namespace KPLN_Library_SQLiteWorker.Core.SQLiteData
{
    /// <summary>
    /// Класс отдела KPLN
    /// </summary>
    public class DBSubDepartment : IDBEntity
    {
        #region Столбцы из БД
        [Key]
        public int Id { get; set; }

        /// <summary>
        /// Код отдела
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// Имя отдела
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Отображение влк/выкл (True/False) для окна авторизации. В БД тип данных текст, преобразование происходит в Dapper
        /// </summary>
        public bool IsAuthEnabled { get; set; }
        #endregion

        /// <summary>
        /// Привязка к БД из DB_Enumerator
        /// </summary>
        public static DB_Enumerator CurrentDB { get; } = DB_Enumerator.SubDepartments;

    }
}
