using System;
using Newtonsoft.Json.Linq;

namespace KPLN_CoordiantorAI.ExternalAIModel.Mcp
{
    internal sealed class RevitMcpToolArgumentException : Exception
    {
        public RevitMcpToolArgumentException(string message, JObject details)
            : base(message)
        {
            Details = details ?? new JObject();
        }

        public JObject Details { get; private set; }
    }
}
