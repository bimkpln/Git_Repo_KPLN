using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using KPLN_CoordiantorAI.ExternalAIModel.Mcp;
using Newtonsoft.Json.Linq;

namespace KPLN_CoordiantorAI.ExternalAIModel.McpClient
{
    internal sealed class AutodeskProductHelpMcpClient : IDisposable
    {
        public const string Endpoint = "https://developer.api.autodesk.com/knowledge/public/v1/mcp";

        private const string GetAvailableProductsTool = "get_available_products";
        private const string SearchHelpContentTool = "search_help_content";
        private const int DefaultSearchResultCount = 3;
        private const int MaximumSearchResultCount = 5;

        private static readonly HashSet<string> ToolNames = new HashSet<string>(StringComparer.Ordinal)
        {
            GetAvailableProductsTool,
            SearchHelpContentTool
        };

        private readonly InternalMcpClient _client;
        private readonly int _revitVersion;

        public AutodeskProductHelpMcpClient(int revitVersion)
        {
            _revitVersion = revitVersion;
            _client = new InternalMcpClient(Endpoint, null, TimeSpan.FromSeconds(90));
        }

        public static bool IsProductHelpTool(string toolName)
        {
            return !string.IsNullOrWhiteSpace(toolName) && ToolNames.Contains(toolName);
        }

        public static IReadOnlyList<RevitMcpToolDefinition> GetHistoryToolDefinitions()
        {
            return new List<RevitMcpToolDefinition>
            {
                new RevitMcpToolDefinition
                {
                    Name = GetAvailableProductsTool,
                    Description = "Returns the official Autodesk product catalog and supported release codes.",
                    InputSchema = RevitMcpToolDefinition.CreateEmptyInputSchema(),
                    RiskLevel = RevitMcpToolRiskLevel.ReadOnly,
                    Area = "AutodeskHelp"
                },
                new RevitMcpToolDefinition
                {
                    Name = SearchHelpContentTool,
                    Description = "Searches official Autodesk Help and Support content for Revit reference questions.",
                    InputSchema = RevitMcpToolDefinition.CreateEmptyInputSchema(),
                    RiskLevel = RevitMcpToolRiskLevel.ReadOnly,
                    Area = "AutodeskHelp"
                }
            };
        }

        public async Task<JArray> GetOpenAiCompatibleToolsAsync(CancellationToken cancellationToken)
        {
            InternalMcpToolProvider provider = new InternalMcpToolProvider(_client);
            JArray tools = await provider.GetOpenAiCompatibleToolsAsync(cancellationToken);

            foreach (JObject tool in tools)
            {
                JObject function = tool["function"] as JObject;
                if (function == null)
                    continue;

                string toolName = function["name"] == null ? string.Empty : function["name"].ToString();
                if (!IsProductHelpTool(toolName))
                    continue;

                string description = function["description"] == null
                    ? string.Empty
                    : function["description"].ToString();
                function["description"] = "Official read-only Autodesk Product Help MCP. " + description;

                if (string.Equals(toolName, SearchHelpContentTool, StringComparison.Ordinal))
                    LimitSearchResultSchema(function["parameters"] as JObject);
            }

            return tools;
        }

        public Task<JObject> CallToolAsync(
            string toolName,
            JObject arguments,
            CancellationToken cancellationToken)
        {
            if (!IsProductHelpTool(toolName))
                throw new InvalidOperationException("Unsupported Autodesk Product Help MCP tool: " + toolName);

            JObject normalizedArguments = arguments == null
                ? new JObject()
                : (JObject)arguments.DeepClone();

            if (string.Equals(toolName, SearchHelpContentTool, StringComparison.Ordinal))
                ApplyRevitSearchDefaults(normalizedArguments);

            return _client.CallToolAsync(toolName, normalizedArguments, cancellationToken);
        }

        public void Dispose()
        {
            _client.Dispose();
        }

        private void ApplyRevitSearchDefaults(JObject arguments)
        {
            if (arguments["locale"] == null || string.IsNullOrWhiteSpace(arguments["locale"].ToString()))
                arguments["locale"] = "en_US";

            if ((arguments["product_code"] == null || string.IsNullOrWhiteSpace(arguments["product_code"].ToString()))
                && (arguments["product_name"] == null || string.IsNullOrWhiteSpace(arguments["product_name"].ToString())))
            {
                arguments["product_code"] = "RVT";
            }

            if (_revitVersion > 0
                && (arguments["release_code"] == null || string.IsNullOrWhiteSpace(arguments["release_code"].ToString())))
            {
                arguments["release_code"] = _revitVersion.ToString(CultureInfo.InvariantCulture);
            }

            int requestedResults = DefaultSearchResultCount;
            JToken maxResultsToken = arguments["max_results"];
            if (maxResultsToken != null && maxResultsToken.Type == JTokenType.Integer)
                requestedResults = maxResultsToken.Value<int>();

            arguments["max_results"] = Math.Max(1, Math.Min(MaximumSearchResultCount, requestedResults));
        }

        private static void LimitSearchResultSchema(JObject schema)
        {
            JObject properties = schema == null ? null : schema["properties"] as JObject;
            JObject maxResults = properties == null ? null : properties["max_results"] as JObject;
            if (maxResults == null)
                return;

            maxResults["default"] = DefaultSearchResultCount;
            JArray alternatives = maxResults["anyOf"] as JArray;
            if (alternatives == null)
                return;

            foreach (JObject alternative in alternatives)
            {
                if (alternative["type"] != null
                    && string.Equals(alternative["type"].ToString(), "integer", StringComparison.Ordinal))
                {
                    alternative["maximum"] = MaximumSearchResultCount;
                }
            }
        }
    }
}
