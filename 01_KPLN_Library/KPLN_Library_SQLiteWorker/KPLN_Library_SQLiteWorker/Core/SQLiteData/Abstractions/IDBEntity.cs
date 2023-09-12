using System.ComponentModel.DataAnnotations;

namespace KPLN_Library_SQLiteWorker.Core.SQLiteData.Abstractions
{
    public interface IDBEntity
    {
        /// <summary>
        /// Id элемента в БД
        /// </summary>
        [Key]
        int Id { get; set; }

        /// <summary>
        /// Привязка к БД из DB_Enumerator
        /// </summary>
        DB_Enumerator CurrentDB { get; set; }
    }
}
