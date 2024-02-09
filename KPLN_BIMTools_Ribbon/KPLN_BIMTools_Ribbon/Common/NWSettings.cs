using Autodesk.Revit.DB;

namespace KPLN_BIMTools_Ribbon.Common
{
    /// <summary>
    /// Настройки экспорта в Navisworks
    /// </summary>
    internal class NWSettings
    {
        /// <summary>
        /// Коэф. фасетизации
        /// </summary>
        internal double FacetingFactor { get; set; }

        /// <summary>
        /// Преобразовать свойства объектов
        /// </summary>
        internal bool ConvertElementProperties { get; set; }

        /// <summary>
        /// Преобразовать связанные файлы
        /// </summary>
        internal bool ExportLinks { get; set; }

        /// <summary>
        /// Проверять и находить отсутсв. материалы
        /// </summary>
        internal bool FindMissingMaterials { get; set; }

        /// <summary>
        /// Область экспорта (файл, вид, выбранные эл-ты)
        /// </summary>
        internal NavisworksExportScope ExportScope { get; set; }

        /// <summary>
        /// Разделять файл по уровням
        /// </summary>
        internal bool DivideFileIntoLevels { get; set; }

        /// <summary>
        /// Экспортировать геомтерию помещений
        /// </summary>
        internal bool ExportRoomGeometry { get; set; }
    }
}
