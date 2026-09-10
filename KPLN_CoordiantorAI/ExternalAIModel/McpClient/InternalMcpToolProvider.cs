using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace KPLN_CoordiantorAI.ExternalAIModel.McpClient
{
    internal sealed class InternalMcpToolProvider
    {
        private readonly InternalMcpClient _client;

        public InternalMcpToolProvider(InternalMcpClient client)
        {
            _client = client;
        }

        public async Task<JArray> GetOpenAiCompatibleToolsAsync(CancellationToken cancellationToken)
        {
            JArray mcpTools = await _client.ListToolsAsync(cancellationToken);
            JArray result = new JArray();

            foreach (JObject tool in mcpTools)
            {
                result.Add(new JObject
                {
                    { "type", "function" },
                    {
                        "function",
                        new JObject
                        {
                            { "name", tool["name"] == null ? string.Empty : tool["name"].ToString() },
                            { "description", tool["description"] == null ? string.Empty : tool["description"].ToString() },
                            { "parameters", tool["inputSchema"] == null ? new JObject() : tool["inputSchema"].DeepClone() }
                        }
                    }
                });
            }

            return result;
        }
    }
}
