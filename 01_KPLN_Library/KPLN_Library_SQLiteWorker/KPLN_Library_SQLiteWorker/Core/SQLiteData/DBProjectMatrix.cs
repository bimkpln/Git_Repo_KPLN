﻿using KPLN_Library_SQLiteWorker.Core.SQLiteData.Abstractions;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace KPLN_Library_SQLiteWorker.Core.SQLiteData
{
    /// <summary>
    /// Матрица допуска к проектам KPLN
    /// </summary>
    public class DBProjectMatrix : IDBEntity
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
        /// Пользователь, которому открыт доступ к проекту
        /// </summary>
        [ForeignKey(nameof(DBUser))]
        public int UserId { get; set; }
    }
}
