using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace KPLN_Loader.Core.SQLiteData
{
    internal class User
    {
        /// <summary>
        /// Id пользователя
        /// </summary>
        [Key]
        public int Id { get; set; }

        /// <summary>
        /// Системное имя пользователя (Windows)
        /// </summary>
        public string SystemName { get; set; }

        /// <summary>
        /// Имя пользователя
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Фамилия пользователя
        /// </summary>
        public string Surname { get; set; }

        /// <summary>
        /// Отдел пользователя
        /// </summary>
        [ForeignKey(nameof(SubDepartment))]
        public int SubDepartmentId { get; set; }

        /// <summary>
        /// Дата регистрации
        /// </summary>
        public string RegistrationDate { get; set; }

        /// <summary>
        /// Дата подключения
        /// </summary>
        public string LastConnectionDate { get; set; }
        
        /// <summary>
        /// Имя Revit-пользователя 
        /// </summary>
        public string RevitUserName { get; set; }

        /// <summary>
        /// Режим отладки вкл/выкл (True/False)
        /// </summary>
        public string IsDebugMode { get; set; }
    }
}
