using System.Collections.Generic;

namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    internal sealed class ParameterDependency
    {
        internal string Name, Formula, SourceValue;
        internal int? IntegerValue;
        internal bool IsReadOnly, IsReporting, IsInstance;
        internal List<string> Inputs = new List<string>(), Dependents = new List<string>(), Links = new List<string>();
    }
}
