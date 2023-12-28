using KPLN_Library_SQLiteWorker.Core.SQLiteData.Abstractions;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace KPLN_Library_SQLiteWorker.Core.SQLiteData
{
    /// <summary>
    /// Класс проекта KPLN
    /// </summary>
    public class DBSubDepartmentCodeMatrix : IDBEntity
    {
        [Key]
        public int Id { get; set; }

        public DB_Enumerator CurrentDB { get; set; }

        /// <summary>
        /// Код (аббревиатура) раздела
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// Id раздела
        /// </summary>
        [ForeignKey(nameof(DBSubDepartment))]
        public int SubDepartmentId { get; set; }
    }
}
