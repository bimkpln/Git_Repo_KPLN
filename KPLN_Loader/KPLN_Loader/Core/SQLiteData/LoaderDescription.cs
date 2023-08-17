using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Loader.Core.SQLiteData
{
    internal class LoaderDescription
    {
        /// <summary>
        /// Id подсказки
        /// </summary>
        [Key]
        public int Id { get; set; }

        /// <summary>
        /// Тело подсказки
        /// </summary>
        public string Description { get; set; }
        
        /// <summary>
        /// Ссылка подсказки
        /// </summary>
        public string InstructionURL { get; set; }

        /// <summary>
        /// Id отдела, для которого подходят подсказки
        /// </summary>
        [ForeignKey(nameof(SubDepartment))]
        public int SubDepartmentId { get; set; }
    }
}
