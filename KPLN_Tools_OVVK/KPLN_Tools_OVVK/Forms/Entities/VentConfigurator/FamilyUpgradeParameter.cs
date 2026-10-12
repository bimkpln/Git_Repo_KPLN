namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    internal sealed class FamilyUpgradeParameter
    {
        internal string Kind, Formula;
        internal bool ReadOnly, Reporting;

        internal bool Writable
        {
            get
            {
                return !ReadOnly && !Reporting && string.IsNullOrWhiteSpace(Formula);
            }
        }

        internal bool SameDefinition(FamilyUpgradeParameter other, bool valve)
        {
            return Kind == other.Kind && ReadOnly == other.ReadOnly && Reporting == other.Reporting
                && (Formula == other.Formula || valve && ValveControl.ConstantFormula(Formula).HasValue && other.Writable);
        }
    }
}
