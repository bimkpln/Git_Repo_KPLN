using static KPLN_Tools_OVVK.Services.VentConfigurator.VentConfiguratorSettings;

namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    internal sealed class ProjectItem
    {
        internal string Name
        { get; set; }

        internal string Stage
        { get; set; }

        internal string Code
        { get; set; }

        internal string[] Paths
        { get; set; }

        internal bool IsUnknown
        { get; set; }

        public string DisplayName
        {
            get
            {
                return IsUnknown ? UnknownProjectName : Name + " [" + Stage + "]";
            }
        }

        public string Description
        {
            get
            {
                return DisplayName + (Paths == null || Paths.Length == 0
                ? string.Empty : "\n" + string.Join("\n", Paths));
            }
        }
    }
}
