using KPLN_Library_SQLiteWorker.Core.SQLiteData.Abstractions;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace KPLN_Library_SQLiteWorker.Core.SQLiteData
{
    /// <summary>
    /// Класс файла для обмена
    /// </summary>
    public class DBRevitDocExchanges : IDBEntity
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
        /// Тип обмена файлов
        /// </summary>
        public string RevitDocExchangeType { get; set; }

        /// <summary>
        /// Имя конфига
        /// </summary>
        public string SettingName { get; set; }

        /// <summary>
        /// Путь к месту сохранения результатов конфига
        /// </summary>
        public string SettingResultPath { get; set; }

        /// <summary>
        /// Количество элементов, которые конфиг обработает
        /// </summary>
        public int SettingCountItem { get; set; }

        /// <summary>
        /// Путь к файлу конфигураций по обмену
        /// </summary>
        public string SettingDBFilePath { get; set; }

        /// <summary>
        /// (УДАЛИТЬ!!! в том числе из БД - оставил архивом, чтобы плагин не падал)
        /// Режим блокировки файла под действия по обмену (True/False). В БД тип данных текст, преобразование происходит в Dapper
        /// </summary>
        public bool IsActive { get; set; }
        #endregion

        /// <summary>
        /// Привязка к БД из DB_Enumerator
        /// </summary>
        public static DB_Enumerator CurrentDB { get; } = DB_Enumerator.RevitDocExchanges;
    }

    public enum RevitDocExchangeEnum
    {
        Navisworks,
        Revit
    }
}
