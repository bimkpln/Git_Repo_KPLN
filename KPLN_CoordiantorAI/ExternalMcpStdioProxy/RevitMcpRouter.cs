using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace RevitMcpStdioProxy
{
    internal sealed class RevitMcpRouter
    {
        private const string ProtocolVersion = "2025-06-18";
        private const string LegacyEndpoint = "http://127.0.0.1:48731/mcp/";
        private const string ListInstancesTool = "list_revit_instances";
        private const string SelectInstanceTool = "select_revit_instance";
        private const string GetSelectedInstanceTool = "get_selected_revit_instance";

        private readonly HttpClient _httpClient;
        private readonly ProxyOptions _options;
        private readonly RevitInstanceDiscovery _discovery;
        private readonly JavaScriptSerializer _serializer;
        private string _selectedInstanceId;

        public RevitMcpRouter(HttpClient httpClient, ProxyOptions options)
        {
            _httpClient = httpClient;
            _options = options;
            _discovery = new RevitInstanceDiscovery();
            _serializer = new JavaScriptSerializer();
        }

        public async Task<string> HandleAsync(string json)
        {
            if (!_options.UsesDiscovery)
                return await ForwardToRevitAsync(_options.DirectEndpoint, json);

            object parsed;
            try
            {
                parsed = _serializer.DeserializeObject(json);
            }
            catch (Exception ex)
            {
                return CreateJsonRpcError("null", -32700, "Invalid JSON: " + ex.Message);
            }

            Dictionary<string, object> request = parsed as Dictionary<string, object>;
            if (request == null)
            {
                string batchErrorCode;
                string batchErrorMessage;
                RevitInstanceInfo batchTarget = ResolveToolTarget(
                    out batchErrorCode,
                    out batchErrorMessage);
                return batchTarget == null
                    ? CreateJsonRpcError("null", -32000, batchErrorMessage)
                    : await ForwardToRevitAsync(batchTarget.Endpoint, json);
            }

            string idJson = GetIdJson(request);
            string method = GetString(request, "method");

            if (string.Equals(method, "initialize", StringComparison.Ordinal))
                return CreateInitializeResponse(idJson, request);

            if (string.Equals(method, "ping", StringComparison.Ordinal)
                || string.Equals(method, "notifications/initialized", StringComparison.Ordinal))
            {
                return CreateJsonRpcResult(idJson, new Dictionary<string, object>());
            }

            if (string.Equals(method, "tools/list", StringComparison.Ordinal))
                return await CreateToolsListResponseAsync(idJson, json);

            if (string.Equals(method, "tools/call", StringComparison.Ordinal))
                return await HandleToolCallAsync(idJson, request, json);

            string errorCode;
            string errorMessage;
            RevitInstanceInfo target = ResolveToolTarget(out errorCode, out errorMessage);
            return target == null
                ? CreateJsonRpcError(idJson, -32000, errorMessage)
                : await ForwardToRevitAsync(target.Endpoint, json);
        }

        private async Task<string> HandleToolCallAsync(
            string idJson,
            Dictionary<string, object> request,
            string originalJson)
        {
            Dictionary<string, object> parameters = GetDictionary(request, "params");
            string toolName = GetString(parameters, "name");
            Dictionary<string, object> arguments = GetDictionary(parameters, "arguments");

            if (string.Equals(toolName, ListInstancesTool, StringComparison.Ordinal))
                return CreateListInstancesResponse(idJson);

            if (string.Equals(toolName, SelectInstanceTool, StringComparison.Ordinal))
                return CreateSelectInstanceResponse(idJson, arguments);

            if (string.Equals(toolName, GetSelectedInstanceTool, StringComparison.Ordinal))
                return CreateGetSelectedInstanceResponse(idJson);

            string errorCode;
            string errorMessage;
            RevitInstanceInfo target = ResolveToolTarget(out errorCode, out errorMessage);
            if (target == null)
            {
                if (CanUseLegacyEndpoint(errorCode))
                {
                    try
                    {
                        return await ForwardToRevitAsync(LegacyEndpoint, originalJson);
                    }
                    catch
                    {
                    }
                }

                return CreateToolErrorResponse(
                    idJson,
                    errorCode,
                    errorMessage,
                    CreateInstancesPayload(_discovery.Discover()));
            }

            return await ForwardToRevitAsync(target.Endpoint, originalJson);
        }

        private string CreateListInstancesResponse(string idJson)
        {
            List<RevitInstanceInfo> instances = _discovery.Discover();
            TryAutoSelect(instances);

            Dictionary<string, object> payload = CreateInstancesPayload(instances);
            payload["selection_required"] =
                string.IsNullOrWhiteSpace(_selectedInstanceId)
                && ApplyConfiguredSelectors(instances).Count > 1;

            return CreateToolSuccessResponse(idJson, payload);
        }

        private string CreateSelectInstanceResponse(
            string idJson,
            Dictionary<string, object> arguments)
        {
            string requestedInstanceId = GetString(arguments, "instance_id");
            if (string.IsNullOrWhiteSpace(requestedInstanceId))
            {
                return CreateToolErrorResponse(
                    idJson,
                    "REVIT_INSTANCE_ID_REQUIRED",
                    "Argument instance_id is required.",
                    CreateInstancesPayload(_discovery.Discover()));
            }

            List<RevitInstanceInfo> allInstances = _discovery.Discover();
            List<RevitInstanceInfo> allowedInstances = ApplyConfiguredSelectors(allInstances);
            RevitInstanceInfo selected = allowedInstances.FirstOrDefault(
                i => string.Equals(
                    i.InstanceId,
                    requestedInstanceId,
                    StringComparison.OrdinalIgnoreCase));

            if (selected == null)
            {
                return CreateToolErrorResponse(
                    idJson,
                    "REVIT_INSTANCE_NOT_FOUND",
                    "The requested Revit instance is not running or is excluded by proxy selectors.",
                    CreateInstancesPayload(allInstances));
            }

            _selectedInstanceId = selected.InstanceId;
            Dictionary<string, object> payload = new Dictionary<string, object>
            {
                { "selected", true },
                { "instance", CreateInstancePayload(selected, true, true) },
                {
                    "message",
                    "All subsequent Revit tool calls in this MCP connection will be routed to this instance."
                }
            };

            return CreateToolSuccessResponse(idJson, payload);
        }

        private string CreateGetSelectedInstanceResponse(string idJson)
        {
            List<RevitInstanceInfo> instances = _discovery.Discover();
            TryAutoSelect(instances);

            if (string.IsNullOrWhiteSpace(_selectedInstanceId))
            {
                Dictionary<string, object> emptyPayload = CreateInstancesPayload(instances);
                emptyPayload["selected"] = false;
                return CreateToolSuccessResponse(idJson, emptyPayload);
            }

            RevitInstanceInfo selected = instances.FirstOrDefault(
                i => string.Equals(
                    i.InstanceId,
                    _selectedInstanceId,
                    StringComparison.OrdinalIgnoreCase));
            if (selected == null)
            {
                return CreateToolErrorResponse(
                    idJson,
                    "REVIT_INSTANCE_UNAVAILABLE",
                    "The previously selected Revit instance is no longer running. Select another instance explicitly.",
                    CreateInstancesPayload(instances));
            }

            return CreateToolSuccessResponse(
                idJson,
                new Dictionary<string, object>
                {
                    { "selected", true },
                    { "instance", CreateInstancePayload(selected, true, true) }
                });
        }

        private async Task<string> CreateToolsListResponseAsync(string idJson, string originalJson)
        {
            RevitInstanceInfo schemaTarget = ResolveSchemaTarget();
            if (schemaTarget == null)
            {
                if (!_options.HasDiscoverySelectors)
                {
                    try
                    {
                        string legacyResponse = await ForwardToRevitAsync(
                            LegacyEndpoint,
                            originalJson);
                        return AppendControlTools(legacyResponse);
                    }
                    catch
                    {
                    }
                }

                return CreateJsonRpcResult(
                    idJson,
                    new Dictionary<string, object>
                    {
                        { "tools", CreateControlToolDefinitions().ToArray() }
                    });
            }

            string responseJson = await ForwardToRevitAsync(schemaTarget.Endpoint, originalJson);
            return AppendControlTools(responseJson);
        }

        private RevitInstanceInfo ResolveToolTarget(
            out string errorCode,
            out string errorMessage)
        {
            List<RevitInstanceInfo> instances = _discovery.Discover();
            if (!string.IsNullOrWhiteSpace(_selectedInstanceId))
            {
                RevitInstanceInfo selected = instances.FirstOrDefault(
                    i => string.Equals(
                        i.InstanceId,
                        _selectedInstanceId,
                        StringComparison.OrdinalIgnoreCase));
                if (selected != null)
                {
                    errorCode = null;
                    errorMessage = null;
                    return selected;
                }

                errorCode = "REVIT_INSTANCE_UNAVAILABLE";
                errorMessage =
                    "The previously selected Revit instance is no longer running. "
                    + "Call list_revit_instances and select_revit_instance.";
                return null;
            }

            List<RevitInstanceInfo> candidates = ApplyConfiguredSelectors(instances);
            if (candidates.Count == 1)
            {
                _selectedInstanceId = candidates[0].InstanceId;
                errorCode = null;
                errorMessage = null;
                return candidates[0];
            }

            if (candidates.Count == 0)
            {
                errorCode = "REVIT_INSTANCE_NOT_FOUND";
                errorMessage =
                    "No matching Revit MCP instance is running. "
                    + "Start Revit with the Coordinator AI plugin loaded.";
                return null;
            }

            errorCode = "REVIT_INSTANCE_REQUIRED";
            errorMessage =
                "Multiple Revit instances are running. "
                + "Call list_revit_instances and then select_revit_instance before using Revit tools.";
            return null;
        }

        private RevitInstanceInfo ResolveSchemaTarget()
        {
            List<RevitInstanceInfo> instances = _discovery.Discover();
            if (!string.IsNullOrWhiteSpace(_selectedInstanceId))
            {
                RevitInstanceInfo selected = instances.FirstOrDefault(
                    i => string.Equals(
                        i.InstanceId,
                        _selectedInstanceId,
                        StringComparison.OrdinalIgnoreCase));
                if (selected != null)
                    return selected;
            }

            List<RevitInstanceInfo> candidates = ApplyConfiguredSelectors(instances);
            if (candidates.Count == 1)
            {
                _selectedInstanceId = candidates[0].InstanceId;
                return candidates[0];
            }

            return candidates.FirstOrDefault();
        }

        private void TryAutoSelect(List<RevitInstanceInfo> instances)
        {
            if (!string.IsNullOrWhiteSpace(_selectedInstanceId))
                return;

            List<RevitInstanceInfo> candidates = ApplyConfiguredSelectors(instances);
            if (candidates.Count == 1)
                _selectedInstanceId = candidates[0].InstanceId;
        }

        private List<RevitInstanceInfo> ApplyConfiguredSelectors(
            IEnumerable<RevitInstanceInfo> instances)
        {
            IEnumerable<RevitInstanceInfo> result = instances
                ?? Enumerable.Empty<RevitInstanceInfo>();

            if (!string.IsNullOrWhiteSpace(_options.InstanceId))
            {
                result = result.Where(i => string.Equals(
                    i.InstanceId,
                    _options.InstanceId,
                    StringComparison.OrdinalIgnoreCase));
            }

            if (_options.ProcessId.HasValue)
                result = result.Where(i => i.ProcessId == _options.ProcessId.Value);

            if (_options.RevitVersion.HasValue)
                result = result.Where(i => i.RevitVersion == _options.RevitVersion.Value);

            if (!string.IsNullOrWhiteSpace(_options.ModelName))
            {
                result = result.Where(
                    i => ContainsIgnoreCase(i.ModelName, _options.ModelName)
                        || ContainsIgnoreCase(i.WindowTitle, _options.ModelName));
            }

            return result.ToList();
        }

        private Dictionary<string, object> CreateInstancesPayload(
            List<RevitInstanceInfo> instances)
        {
            List<RevitInstanceInfo> matchingInstances = ApplyConfiguredSelectors(instances);
            HashSet<string> matchingIds = new HashSet<string>(
                matchingInstances.Select(i => i.InstanceId),
                StringComparer.OrdinalIgnoreCase);

            List<object> values = instances
                .Select(
                    i => (object)CreateInstancePayload(
                        i,
                        string.Equals(
                            i.InstanceId,
                            _selectedInstanceId,
                            StringComparison.OrdinalIgnoreCase),
                        matchingIds.Contains(i.InstanceId)))
                .ToList();

            return new Dictionary<string, object>
            {
                { "instances", values.ToArray() },
                { "count", values.Count },
                { "selected_instance_id", _selectedInstanceId },
                {
                    "selected_instance_available",
                    !string.IsNullOrWhiteSpace(_selectedInstanceId)
                    && instances.Any(i => string.Equals(
                        i.InstanceId,
                        _selectedInstanceId,
                        StringComparison.OrdinalIgnoreCase))
                }
            };
        }

        private static Dictionary<string, object> CreateInstancePayload(
            RevitInstanceInfo instance,
            bool selected,
            bool matchesSelectors)
        {
            return new Dictionary<string, object>
            {
                { "instance_id", instance.InstanceId },
                { "process_id", instance.ProcessId },
                { "revit_version", instance.RevitVersion },
                { "model_name", instance.ModelName },
                { "window_title", instance.WindowTitle },
                { "endpoint", instance.Endpoint },
                { "plugin_version", instance.PluginVersion },
                { "process_start_time_utc", instance.ProcessStartTimeUtc.ToString("o") },
                { "selected", selected },
                { "matches_proxy_selectors", matchesSelectors }
            };
        }

        private string AppendControlTools(string responseJson)
        {
            Dictionary<string, object> response;
            try
            {
                response = _serializer.DeserializeObject(responseJson) as Dictionary<string, object>;
            }
            catch
            {
                return responseJson;
            }

            if (response == null)
                return responseJson;

            Dictionary<string, object> result = GetDictionary(response, "result");
            if (result.Count == 0)
                return responseJson;

            List<object> tools = new List<object>();
            object existingTools;
            if (result.TryGetValue("tools", out existingTools))
            {
                IEnumerable enumerable = existingTools as IEnumerable;
                if (enumerable != null && !(existingTools is string))
                {
                    foreach (object tool in enumerable)
                        tools.Add(tool);
                }
            }

            foreach (Dictionary<string, object> controlTool in CreateControlToolDefinitions())
            {
                string controlName = GetString(controlTool, "name");
                if (!tools.Any(t => string.Equals(
                    GetString(t as Dictionary<string, object>, "name"),
                    controlName,
                    StringComparison.Ordinal)))
                {
                    tools.Add(controlTool);
                }
            }

            result["tools"] = tools.ToArray();
            response["result"] = result;
            return _serializer.Serialize(response);
        }

        private static List<Dictionary<string, object>> CreateControlToolDefinitions()
        {
            Dictionary<string, object> emptySchema = new Dictionary<string, object>
            {
                { "type", "object" },
                { "properties", new Dictionary<string, object>() },
                { "additionalProperties", false }
            };

            Dictionary<string, object> selectSchema = new Dictionary<string, object>
            {
                { "type", "object" },
                {
                    "properties",
                    new Dictionary<string, object>
                    {
                        {
                            "instance_id",
                            new Dictionary<string, object>
                            {
                                { "type", "string" },
                                {
                                    "description",
                                    "Exact instance_id returned by list_revit_instances."
                                }
                            }
                        }
                    }
                },
                { "required", new[] { "instance_id" } },
                { "additionalProperties", false }
            };

            return new List<Dictionary<string, object>>
            {
                CreateToolDefinition(
                    ListInstancesTool,
                    "Lists all live local Revit processes registered by Coordinator AI. "
                    + "When more than one process is running, call this tool and then "
                    + "select_revit_instance before calling Revit tools.",
                    emptySchema),
                CreateToolDefinition(
                    SelectInstanceTool,
                    "Selects the exact Revit process used by subsequent Revit tool calls "
                    + "in this MCP connection. The proxy never silently switches away "
                    + "from a selected process.",
                    selectSchema),
                CreateToolDefinition(
                    GetSelectedInstanceTool,
                    "Returns the Revit process currently selected for this MCP connection.",
                    emptySchema)
            };
        }

        private static Dictionary<string, object> CreateToolDefinition(
            string name,
            string description,
            Dictionary<string, object> inputSchema)
        {
            return new Dictionary<string, object>
            {
                { "name", name },
                { "description", description },
                { "inputSchema", inputSchema }
            };
        }

        private string CreateInitializeResponse(
            string idJson,
            Dictionary<string, object> request)
        {
            Dictionary<string, object> parameters = GetDictionary(request, "params");
            string requestedVersion = GetString(parameters, "protocolVersion");
            if (string.IsNullOrWhiteSpace(requestedVersion))
                requestedVersion = ProtocolVersion;

            return CreateJsonRpcResult(
                idJson,
                new Dictionary<string, object>
                {
                    { "protocolVersion", requestedVersion },
                    {
                        "capabilities",
                        new Dictionary<string, object>
                        {
                            { "tools", new Dictionary<string, object>() }
                        }
                    },
                    {
                        "serverInfo",
                        new Dictionary<string, object>
                        {
                            { "name", "revit-mcp-router" },
                            { "version", "0.2.0" }
                        }
                    }
                });
        }

        private string CreateToolSuccessResponse(
            string idJson,
            Dictionary<string, object> payload)
        {
            string text = _serializer.Serialize(payload);
            Dictionary<string, object> result = new Dictionary<string, object>
            {
                {
                    "content",
                    new object[]
                    {
                        new Dictionary<string, object>
                        {
                            { "type", "text" },
                            { "text", text }
                        }
                    }
                },
                { "isError", false },
                { "structuredContent", payload }
            };

            return CreateJsonRpcResult(idJson, result);
        }

        private string CreateToolErrorResponse(
            string idJson,
            string code,
            string message,
            Dictionary<string, object> context)
        {
            Dictionary<string, object> error = new Dictionary<string, object>
            {
                { "code", code },
                { "message", message }
            };
            if (context != null)
                error["context"] = context;

            Dictionary<string, object> structuredContent = new Dictionary<string, object>
            {
                { "success", false },
                { "error", error }
            };

            Dictionary<string, object> result = new Dictionary<string, object>
            {
                {
                    "content",
                    new object[]
                    {
                        new Dictionary<string, object>
                        {
                            { "type", "text" },
                            { "text", code + ": " + message }
                        }
                    }
                },
                { "isError", true },
                { "structuredContent", structuredContent }
            };

            return CreateJsonRpcResult(idJson, result);
        }

        private async Task<string> ForwardToRevitAsync(string endpoint, string json)
        {
            using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, endpoint))
            {
                request.Content = new StringContent(json, Encoding.UTF8, "application/json");
                request.Headers.Add("X-Revit-MCP-Client", ResolveClientName());

                using (HttpResponseMessage response = await _httpClient.SendAsync(request))
                {
                    string responseJson = await response.Content.ReadAsStringAsync();
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new InvalidOperationException(
                            "HTTP " + (int)response.StatusCode + ": " + responseJson);
                    }

                    return responseJson;
                }
            }
        }

        private string CreateJsonRpcResult(string idJson, object result)
        {
            return "{\"jsonrpc\":\"2.0\",\"id\":"
                + NormalizeIdJson(idJson)
                + ",\"result\":"
                + _serializer.Serialize(result)
                + "}";
        }

        private string CreateJsonRpcError(string idJson, int code, string message)
        {
            return "{\"jsonrpc\":\"2.0\",\"id\":"
                + NormalizeIdJson(idJson)
                + ",\"error\":{\"code\":"
                + code
                + ",\"message\":"
                + _serializer.Serialize(message)
                + "}}";
        }

        private string GetIdJson(Dictionary<string, object> request)
        {
            object id;
            return request != null && request.TryGetValue("id", out id)
                ? _serializer.Serialize(id)
                : "null";
        }

        private static string NormalizeIdJson(string idJson)
        {
            return string.IsNullOrWhiteSpace(idJson) ? "null" : idJson;
        }

        private static Dictionary<string, object> GetDictionary(
            Dictionary<string, object> source,
            string name)
        {
            if (source == null)
                return new Dictionary<string, object>();

            object value;
            Dictionary<string, object> dictionary;
            return source.TryGetValue(name, out value)
                && (dictionary = value as Dictionary<string, object>) != null
                ? dictionary
                : new Dictionary<string, object>();
        }

        private static string GetString(
            Dictionary<string, object> source,
            string name)
        {
            if (source == null)
                return string.Empty;

            object value;
            return source.TryGetValue(name, out value) && value != null
                ? value.ToString()
                : string.Empty;
        }

        private static bool ContainsIgnoreCase(string value, string search)
        {
            return !string.IsNullOrWhiteSpace(value)
                && !string.IsNullOrWhiteSpace(search)
                && value.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private bool CanUseLegacyEndpoint(string errorCode)
        {
            return string.IsNullOrWhiteSpace(_selectedInstanceId)
                && !_options.HasDiscoverySelectors
                && string.Equals(
                    errorCode,
                    "REVIT_INSTANCE_NOT_FOUND",
                    StringComparison.Ordinal);
        }

        private static string ResolveClientName()
        {
            string value = Environment.GetEnvironmentVariable("REVIT_MCP_CLIENT_NAME");
            return string.IsNullOrWhiteSpace(value) ? "external_mcp" : value.Trim();
        }
    }
}
