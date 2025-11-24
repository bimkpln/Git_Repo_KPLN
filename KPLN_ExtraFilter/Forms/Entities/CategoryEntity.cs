using Autodesk.Revit.DB;

namespace KPLN_ExtraFilter.Forms.Entities
{
    /// <summary>
    /// Класс-сущность для категории
    /// </summary>
    public sealed class CategoryEntity
    {
        public CategoryEntity(Category cat)
        {
            RevitCat = cat;
            RevitCatName = cat.Name;
        }

        public Category RevitCat { get; set; }

        public string RevitCatName { get; set; }
    }
}
