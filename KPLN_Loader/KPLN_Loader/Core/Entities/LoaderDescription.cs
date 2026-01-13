using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace KPLN_Loader.Core.Entities
{
    internal sealed class LoaderDescription
    {
        /// <summary>
        /// Id подсказки
        /// </summary>
        [Key]
        internal int Id { get; set; }

        /// <summary>
        /// Тело подсказки
        /// </summary>
        internal string Description { get; set; }

        /// <summary>
        /// Ссылка подсказки
        /// </summary>
        internal string InstructionURL { get; set; }

        /// <summary>
        /// Id отдела, для которого подходят подсказки
        /// </summary>
        [ForeignKey(nameof(SubDepartment))]
        internal int SubDepartmentId { get; set; }

        /// <summary>
        /// Счётчик "лайков"
        /// </summary>
        internal int ApprovalRate { get; set; }

        /// <summary>
        /// Счётчик "дизлайков"
        /// </summary>
        internal int DisapprovalRate { get; set; }
    }
}
