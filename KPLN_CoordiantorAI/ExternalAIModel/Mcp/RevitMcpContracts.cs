using Newtonsoft.Json.Linq;

namespace KPLN_CoordiantorAI.ExternalAIModel.Mcp
{
    internal enum RevitMcpToolRiskLevel
    {
        ReadOnly,
        UiChanging,
        ModelChanging
    }

    internal class RevitMcpToolDefinition
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public JObject InputSchema { get; set; }
        public RevitMcpToolRiskLevel RiskLevel { get; set; }
        public string Area { get; set; }

        public JObject ToMcpToolObject()
        {
            return new JObject
            {
                { "name", Name },
                { "description", Description },
                { "inputSchema", InputSchema == null ? CreateEmptyInputSchema() : (JObject)InputSchema.DeepClone() }
            };
        }

        public static JObject CreateEmptyInputSchema()
        {
            return new JObject
            {
                { "type", "object" },
                { "properties", new JObject() },
                { "required", new JArray() }
            };
        }
    }

    internal class RevitMcpToolCallRequest
    {
        public string ToolName { get; set; }
        public JObject Arguments { get; set; }
    }

    internal class RevitMcpToolCallResponse
    {
        public bool Success { get; set; }
        public string ToolName { get; set; }
        public JToken Result { get; set; }
        public RevitMcpToolExecutionError Error { get; set; }
        public string RevitModelName { get; set; }
        public string RevitViewName { get; set; }
    }

    internal class RevitMcpToolExecutionError
    {
        public string Code { get; set; }
        public string Message { get; set; }
        public string ExceptionType { get; set; }
        public JObject Details { get; set; }
    }
}
