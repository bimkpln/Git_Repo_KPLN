using KPLN_Library_SQLiteWorker.Core.SQLiteData.Abstractions;
using System.ComponentModel.DataAnnotations;

namespace KPLN_Library_SQLiteWorker.Core.SQLiteData
{
    /// <summary>
    /// Класс всплывающих окон Revit
    /// </summary>
    public class DBRevitDialog : IDBEntity
    {
        #region Столбцы из БД
        [Key]
        public int Id { get; set; }

        /// <summary>
        /// ID диалога (текстовое имя)
        /// </summary>
        public string DialogId { get; set; }

        /// <summary>
        /// Значение, необходимое для закрытия окна
        /// </summary>
        public string OverrideResult { get; set; }
        #endregion

        /// <summary>
        /// Привязка к БД из DB_Enumerator
        /// </summary>
        public static DB_Enumerator CurrentDB { get; } = DB_Enumerator.RevitDialogs;
    }
}
