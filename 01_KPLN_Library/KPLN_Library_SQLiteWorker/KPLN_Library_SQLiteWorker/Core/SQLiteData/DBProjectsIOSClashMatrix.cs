using KPLN_Library_SQLiteWorker.Core.SQLiteData.Abstractions;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace KPLN_Library_SQLiteWorker.Core.SQLiteData
{
    /// <summary>
    /// Строка матрицы допуска к проектам KPLN
    /// </summary>
    public class DBProjectsIOSClashMatrix : IDBEntity
    {
        #region Столбцы из БД
        [Key]
        public int Id { get; set; }

        /// <summary>
        /// Проект, к которому файл относится
        /// </summary>
        [ForeignKey(nameof(DBProject))]
        public int ProjectId { get; set; }

        /// <summary>
        /// Отдел, на пользователей которого правило блокировки НЕ распространяется
        /// </summary>
        [ForeignKey(nameof(DBSubDepartment))]
        public int ExceptionSubDepartmentId { get; set; }

        /// <summary>
        /// Пользователь, на которого правило блокировки НЕ распространяется
        /// </summary>
        [ForeignKey(nameof(DBUser))]
        public int ExceptionUserId { get; set; }
        #endregion

        /// <summary>
        /// Привязка к БД из DB_Enumerator
        /// </summary>
        public static DB_Enumerator CurrentDB { get; } = DB_Enumerator.ProjectsIOSClashMatrix;
    }
}
