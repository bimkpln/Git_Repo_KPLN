using System.ComponentModel.DataAnnotations.Schema;

namespace KPLN_Loader.Core.Entities
{
    /// <summary>
    /// Связь модуля с разделом в таблице ModulesMatrix.
    /// </summary>
    internal sealed class ModuleMatrix
    {
        /// <summary>
        /// Id модуля.
        /// </summary>
        [ForeignKey(nameof(Module))]
        internal int ModuleId { get; set; }

        /// <summary>
        /// Id раздела. Значение 1 разрешает загрузку для всех разделов.
        /// </summary>
        [ForeignKey(nameof(SubDepartment))]
        internal int SubDepartmentId { get; set; }
    }
}
