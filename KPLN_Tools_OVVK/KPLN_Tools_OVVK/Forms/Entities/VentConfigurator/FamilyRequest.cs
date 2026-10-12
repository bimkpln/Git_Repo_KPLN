using System.Collections.Generic;

namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    internal sealed class FamilyRequest
    {
        internal RequestKind Kind
        { get; set; }

        internal List<FamilyTypeItem> Types
        { get; set; }

        internal string ContextKey
        { get; set; }

        internal string TypeName
        { get; set; }

        internal InstallationConfiguration Configuration
        { get; set; }

        internal string FamilyPath
        { get; set; }

        internal string OutputPath
        { get; set; }

        internal bool HasAutomaticName
        { get; set; }

        internal bool IsDirty
        { get; set; }

        internal FamilyTypeItem TargetType
        { get; set; }

        internal string PersistedName
        { get; set; }
    }
}
