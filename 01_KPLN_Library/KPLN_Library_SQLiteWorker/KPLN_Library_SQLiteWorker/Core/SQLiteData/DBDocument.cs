using KPLN_Library_SQLiteWorker.Core.SQLiteData.Abstractions;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace KPLN_Library_SQLiteWorker.Core.SQLiteData
{
    /// <summary>
    /// Класс файла (документа) Revit
    /// </summary>
    public class DBDocument : IDBEntity
    {
        [Key]
        public int Id { get; set; }
        
        public DB_Enumerator CurrentDB { get; set; }

        /// <summary>
        /// Проект, к которому файл относится
        /// </summary>
        [ForeignKey(nameof(DBProject))]
        public int ProjectId { get; set; }

        /// <summary>
        /// Имя файла
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Полный путь к файлу
        /// </summary>
        public string FullPath { get; set; }
    }
}
