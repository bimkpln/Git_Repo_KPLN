using KPLN_Library_ConfigWorker.Core;
using System.Collections.Generic;

namespace KPLN_Tools_OVVK.Forms.Entities
{
    public sealed class OVVK_MergeSystemsConfigM : IJsonSerializable
    {
        public List<ModelSettings> Models { get; set; } = new List<ModelSettings>();

        public object ToJson() => new { Models };

        public sealed class ModelSettings
        {
            public string ModelName { get; set; }
            public string ParameterName { get; set; }
            public List<Group> Groups { get; set; } = new List<Group>();
        }

        public sealed class Group
        {
            public List<string> SystemNames { get; set; }
            public string FirstParamName { get; set; }
            public string LastParamName { get; set; }
        }
    }
}
