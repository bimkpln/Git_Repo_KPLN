using System.ComponentModel.DataAnnotations;

namespace KPLN_Loader.Core.SQLiteData
{
    public class SubDepartment
    {
        /// <summary>
        /// Id отдела
        /// </summary>
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
        /// Отображение влк/выкл (True/False) для окна авторизации
        /// </summary>
        public string IsAuthEnabled { get; set; }
    }
}
