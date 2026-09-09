using System;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using KPLN_CoordiantorAI.ExternalAIModel.Mcp;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace KPLN_CoordiantorAI.ExternalAIModel.McpClient
{
    internal sealed class InternalMcpClient : IDisposable
    {
        private readonly HttpClient _httpClient;
        private readonly string _endpoint;
        private int _nextRequestId = 1;
        private bool _disposed;

        public InternalMcpClient()
            : this(RevitMcpServer.DefaultEndpoint)
        {
        }

        public InternalMcpClient(string endpoint)
        {
            _endpoint = endpoint;
            _httpClient = new HttpClient();
            _httpClient.Timeout = TimeSpan.FromMinutes(5);
        }

        public Task<JObject> InitializeAsync(CancellationToken cancellationToken)
        {
            JObject parameters = new JObject
            {
                { "protocolVersion", "2025-06-18" },
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
                httpRequest.Headers.Add("X-Revit-MCP-Client", "wpf_window");

                using (HttpResponseMessage httpResponse = await _httpClient.SendAsync(httpRequest, cancellationToken))
                {
                    string responseJson = await httpResponse.Content.ReadAsStringAsync();
                    if (!httpResponse.IsSuccessStatusCode)
                        throw new InvalidOperationException("MCP HTTP error " + (int)httpResponse.StatusCode + ": " + responseJson);

                    JObject response = JObject.Parse(responseJson);
                    JObject error = response["error"] as JObject;
                    if (error != null)
                        throw new InvalidOperationException(error["message"] == null ? "MCP error" : error["message"].ToString());

                    return response["result"] as JObject ?? new JObject();
                }
            }
        }
    }
}
