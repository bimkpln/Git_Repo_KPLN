using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace KPLN_CoordiantorAI.ExternalAIModel.McpClient
{
    internal sealed class InternalMcpClient : IDisposable
    {
        private const string ProtocolVersion = "2025-06-18";

        private readonly HttpClient _httpClient;
        private readonly string _endpoint;
        private readonly string _requestScopeId;
        private readonly bool _isLocalEndpoint;
        private int _nextRequestId = 1;
        private bool _disposed;

        public InternalMcpClient()
            : this(ResolveCurrentProcessEndpoint(), null)
        {
        }

        private static string ResolveCurrentProcessEndpoint()
        {
            if (string.IsNullOrWhiteSpace(ModuleData.McpEndpoint))
            {
                throw new InvalidOperationException(
                    "The MCP server for this Revit process is not running. "
                    + "Restart Revit and check the Coordinator AI diagnostic log.");
            }

            return ModuleData.McpEndpoint;
        }

        public InternalMcpClient(string endpoint)
            : this(endpoint, null)
        {
        }

        public InternalMcpClient(string endpoint, string requestScopeId)
            : this(endpoint, requestScopeId, TimeSpan.FromMinutes(5))
        {
        }

        internal InternalMcpClient(string endpoint, string requestScopeId, TimeSpan timeout)
        {
            _endpoint = endpoint;
            _requestScopeId = requestScopeId;
            Uri endpointUri;
            _isLocalEndpoint = Uri.TryCreate(endpoint, UriKind.Absolute, out endpointUri)
                && endpointUri.IsLoopback;
            _httpClient = new HttpClient();
            _httpClient.Timeout = timeout > TimeSpan.Zero
                ? timeout
                : TimeSpan.FromMinutes(5);
        }

        public Task<JObject> InitializeAsync(CancellationToken cancellationToken)
        {
            JObject parameters = new JObject
            {
                { "protocolVersion", ProtocolVersion },
                { "capabilities", new JObject() },
                { "clientInfo", new JObject { { "name", "revit-wpf-chat" }, { "version", "0.1.0" } } }
            };

            return SendRequestAsync("initialize", parameters, cancellationToken);
        }

        public async Task<JArray> ListToolsAsync(CancellationToken cancellationToken)
        {
            JObject response = await SendRequestAsync("tools/list", new JObject(), cancellationToken);
            return response["tools"] as JArray ?? new JArray();
        }

        public Task<JObject> CallToolAsync(string toolName, JObject arguments, CancellationToken cancellationToken)
        {
            JObject parameters = new JObject
            {
                { "name", toolName },
                { "arguments", arguments == null ? new JObject() : arguments }
            };

            return SendRequestAsync("tools/call", parameters, cancellationToken);
        }

        public void Dispose()
        {
            if (_disposed)
                return;

            _httpClient.Dispose();
            _disposed = true;
        }

        private async Task<JObject> SendRequestAsync(
            string method,
            JObject parameters,
            CancellationToken cancellationToken)
        {
            if (_disposed)
                throw new ObjectDisposedException(GetType().FullName);

            JObject request = new JObject
            {
                { "jsonrpc", "2.0" },
                { "id", _nextRequestId++ },
                { "method", method },
                { "params", parameters ?? new JObject() }
            };

            string json = JsonConvert.SerializeObject(request);
            using (HttpRequestMessage httpRequest = new HttpRequestMessage(HttpMethod.Post, _endpoint))
            {
                httpRequest.Content = new StringContent(json, Encoding.UTF8, "application/json");
                httpRequest.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                httpRequest.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("text/event-stream"));
                httpRequest.Headers.TryAddWithoutValidation("MCP-Protocol-Version", ProtocolVersion);
                if (_isLocalEndpoint)
                {
                    httpRequest.Headers.Add("X-Revit-MCP-Client", "wpf_window");
                    if (!string.IsNullOrWhiteSpace(_requestScopeId))
                        httpRequest.Headers.Add("X-Revit-MCP-Request-Scope", _requestScopeId);
                }

                using (HttpResponseMessage httpResponse = await _httpClient.SendAsync(httpRequest, cancellationToken))
                {
                    string responseJson = await httpResponse.Content.ReadAsStringAsync();
                    if (!httpResponse.IsSuccessStatusCode)
                        throw new InvalidOperationException("MCP HTTP error " + (int)httpResponse.StatusCode + ": " + responseJson);

                    JObject response = ParseMcpResponse(responseJson);
                    JObject error = response["error"] as JObject;
                    if (error != null)
                        throw new InvalidOperationException(error["message"] == null ? "MCP error" : error["message"].ToString());

                    return response["result"] as JObject ?? new JObject();
                }
            }
        }

        private static JObject ParseMcpResponse(string responseText)
        {
            if (string.IsNullOrWhiteSpace(responseText))
                throw new InvalidOperationException("MCP server returned an empty response.");

            string trimmed = responseText.Trim();
            if (trimmed.StartsWith("{", StringComparison.Ordinal))
                return JObject.Parse(trimmed);

            JObject lastMessage = null;
            StringBuilder eventData = new StringBuilder();
            string[] lines = responseText.Replace("\r\n", "\n").Split('\n');
            foreach (string line in lines)
            {
                if (line.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
                {
                    if (eventData.Length > 0)
                        eventData.Append('\n');
                    eventData.Append(line.Substring(5).TrimStart());
                    continue;
                }

                if (line.Length == 0 && eventData.Length > 0)
                {
                    JObject parsed = TryParseEventData(eventData.ToString());
                    if (parsed != null)
                        lastMessage = parsed;
                    eventData.Clear();
                }
            }

            if (eventData.Length > 0)
            {
                JObject parsed = TryParseEventData(eventData.ToString());
                if (parsed != null)
                    lastMessage = parsed;
            }

            if (lastMessage == null)
                throw new InvalidOperationException("MCP server returned an unsupported Streamable HTTP response.");

            return lastMessage;
        }

        private static JObject TryParseEventData(string eventData)
        {
            if (string.IsNullOrWhiteSpace(eventData) || string.Equals(eventData.Trim(), "[DONE]", StringComparison.Ordinal))
                return null;

            try
            {
                return JObject.Parse(eventData);
            }
            catch (JsonException)
            {
                return null;
            }
        }
    }
}
