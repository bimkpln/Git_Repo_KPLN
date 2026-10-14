using KPLN_Library_ConfigWorker.Core;
using System.Collections.Generic;

namespace KPLN_Tools_OVVK.Forms.Entities
{
    public sealed class OVVK_MergeSystemsConfigM : IJsonSerializable
    {
        public string DocumentKey { get; set; }
        public string ParameterName { get; set; }
        public List<Group> Groups { get; set; } = new List<Group>();

        public object ToJson() => new { DocumentKey, ParameterName, Groups };

        public sealed class Group
        {
            public string FirstParamName { get; set; }
            public string LastParamName { get; set; }
        }
    }
}
