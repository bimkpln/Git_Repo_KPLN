using KPLN_Library_DBWorker.Core.Abstractions;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace KPLN_Library_DBWorker.Core
{
    /// <summary>
    /// Класс проекта KPLN
    /// </summary>
    public class DBSubDepartmentCodeMatrix : IDBEntity
    {
        #region Столбцы из БД
        [Key]
        public int Id { get; set; }

        /// <summary>
        /// Код (аббревиатура) раздела
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// Id раздела
        /// </summary>
        [ForeignKey(nameof(DBSubDepartment))]
        public int SubDepartmentId { get; set; }
        #endregion

        /// <summary>
        /// Привязка к БД из DB_Enumerator
        /// </summary>
        public static DBEnumerator CurrentDB { get; } = DBEnumerator.SubDepartmentCodeMatrix;
    }
}
