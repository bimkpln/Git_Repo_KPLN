using System.Collections.Generic;

namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    internal sealed class FamilyPackage
    {
        internal SectionCatalog Catalog;
        internal CompositionSnapshot DefinitionSnapshot;
        internal List<FamilyTypeItem> Types = new List<FamilyTypeItem>();
        internal string Path, SourcePath, SelectedTypeName, LoadNotice, UpdateNotice;
    }
}
