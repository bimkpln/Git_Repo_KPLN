using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace KPLN_Loader.Core.Entities
{
    public sealed class User
    {
        /// <summary>
        /// Id пользователя (ключ для БД)
        /// </summary>
        [Key]
        public int Id { get; set; }

        /// <summary>
        /// Системное имя пользователя (Windows)
        /// </summary>
        public string SystemName { get; set; }

        /// <summary>
        /// GUID-пользователя, чтобы исключить возможность наличия одинаковых имен (ключ для Google Tabs)
        /// </summary>
        public string SystemGuid { get; set; }

        /// <summary>
        /// Имя пользователя
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Фамилия пользователя
        /// </summary>
        public string Surname { get; set; }

        /// <summary>
        /// Организация пользователя
        /// </summary>
        public string Company { get; set; }

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
        /// Ограничить работу пользователя вкл/выкл (True/False). В БД тип данных текст, преобразование происходит в Dapper
        /// </summary>
        public bool IsUserRestricted { get; set; }

        /// <summary>
        /// KPLN: Режим отладки вкл/выкл (True/False). В БД тип данных текст, преобразование происходит в Dapper
        /// </summary>
        public bool IsDebugMode { get; set; }

        /// <summary>
        /// KPLN: ID пользователя Bitrix
        /// </summary>
        public int BitrixUserID { get; set; }

        /// <summary>
        /// Прямое указание на то, что пользователь НЕ сотрудник KPLN
        /// </summary>
        public bool IsExtraNet { get; set; }
    }
}
