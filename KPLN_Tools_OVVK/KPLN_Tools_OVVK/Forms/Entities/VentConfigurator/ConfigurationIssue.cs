namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    internal sealed class ConfigurationIssue
    {
        internal string ParameterName, Message, InitialValue;
        internal object SuggestedValue;
        internal bool Accepted;

        internal ConfigurationIssue Copy()
        {
            return (ConfigurationIssue)MemberwiseClone();
        }
    }
}
