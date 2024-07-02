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
        #region Столбцы из БД
        [Key]
        public int Id { get; set; }
        
        /// <summary>
        /// Путь к модели из хранилища
        /// </summary>
        public string CentralPath { get; set; }

        /// <summary>
        /// Проект, к которому файл относится
        /// </summary>
        [ForeignKey(nameof(DBProject))]
        public int ProjectId { get; set; }

        /// <summary>
        /// ID отдела, которому принадлежит файл
        /// </summary>
        [ForeignKey(nameof(DBSubDepartment))]
        public int SubDepartmentId { get; set; }

        /// <summary>
        /// ID специалиста, который последний раз вносил изменения
        /// </summary>
        [ForeignKey(nameof(DBUser))]
        public int LastChangedUserId { get; set; }

        /// <summary>
        /// Дата последнего изменения
        /// </summary>
        public string LastChangedData { get; set; }

        /// <summary>
        /// Метка закрытого проекта. В БД тип данных текст, преобразование происходит в Dapper
        /// </summary>
        public bool IsClosed { get; set; }
        #endregion

        /// <summary>
        /// Привязка к БД из DB_Enumerator
        /// </summary>
        public static DB_Enumerator CurrentDB { get; } = DB_Enumerator.Documents;
    }
}
