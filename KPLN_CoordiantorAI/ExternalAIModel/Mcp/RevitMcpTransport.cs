using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace KPLN_CoordiantorAI.ExternalAIModel.Mcp
{
    internal sealed class RevitMcpTransport
    {
        private const string DefaultProtocolVersion = "2025-06-18";

        private readonly string _prefix;
        private readonly RevitMcpExternalEventHandler _externalEventHandler;
        private HttpListener _listener;
        private CancellationTokenSource _cancellation;
        private Task _listenTask;

        public RevitMcpTransport(string prefix, RevitMcpExternalEventHandler externalEventHandler)
        {
            _prefix = prefix;
            _externalEventHandler = externalEventHandler;
        }

        public bool IsRunning
        {
            get { return _listener != null && _listener.IsListening; }
        }

        public void Start()
        {
            RevitMcpDiagnosticLogger.Log("Transport.Start requested. Prefix=" + _prefix);

            if (IsRunning)
            {
                RevitMcpDiagnosticLogger.Log("Transport.Start skipped: listener already running.");
                return;
            }

            try
            {
                _cancellation = new CancellationTokenSource();
                _listener = new HttpListener();
                _listener.Prefixes.Add(_prefix);
                _listener.Start();

                _listenTask = Task.Run(() => ListenLoop(_cancellation.Token));
                _listenTask.ContinueWith(t =>
                {
                    if (t.Exception != null)
                        RevitMcpDiagnosticLogger.LogException("ListenLoop task faulted", t.Exception);
                }, TaskContinuationOptions.ExecuteSynchronously);
            }
            catch (Exception ex)
            {
                RevitMcpDiagnosticLogger.LogException("Transport.Start failed", ex);
                Stop();
                throw;
            }
        }

        public void Stop()
        {
            RevitMcpDiagnosticLogger.Log("Transport.Stop requested.");

            try
            {
                if (_cancellation != null)
                    _cancellation.Cancel();

                if (_listener != null)
                    _listener.Stop();
            }
            catch
            {
            }
            finally
            {
                if (_listener != null)
                    _listener.Close();

                if (_cancellation != null)
                    _cancellation.Dispose();

                _listener = null;
                _cancellation = null;
                _listenTask = null;
                RevitMcpDiagnosticLogger.Log("Transport.Stop completed.");
            }
        }

        private async Task ListenLoop(CancellationToken cancellationToken)
        {
            while (!cancellationToken.IsCancellationRequested && IsRunning)
            {
                HttpListenerContext context = null;

                try
                {
                    context = await _listener.GetContextAsync();
                    await HandleContextAsync(context, cancellationToken);
                }
                catch (ObjectDisposedException)
                {
                    return;
                }
                catch (HttpListenerException ex)
                {
                    if (!cancellationToken.IsCancellationRequested)
                        RevitMcpDiagnosticLogger.LogException("ListenLoop HttpListenerException", ex);
                    return;
                }
                catch (Exception ex)
                {
                    RevitMcpDiagnosticLogger.LogException("ListenLoop exception", ex);
                    if (context != null)
                        await TryWriteJsonAsync(context.Response, CreateError(null, -32603, ex.Message), 500);
                }
            }
        }

        private async Task HandleContextAsync(HttpListenerContext context, CancellationToken cancellationToken)
        {
            HttpListenerRequest request = context.Request;
            HttpListenerResponse response = context.Response;

            try
            {
                string scenario = ResolveInteractionScenario(request);

                AddCorsHeaders(response);

                if (!RevitMcpSecurity.IsLocalRequest(request))
                {
                    await WriteJsonAsync(response, CreateError(null, -32000, "Revit MCP accepts only local connections."), 403);
                    return;
                }

                if (request.HttpMethod == "OPTIONS")
                {
                    response.StatusCode = 204;
                    return;
                }

                if (request.HttpMethod == "GET")
                {
                    await WriteJsonAsync(response, CreateStatusObject(), 200);
                    return;
                }

                if (request.HttpMethod != "POST")
                {
                    await WriteJsonAsync(response, CreateError(null, -32600, "Only GET, POST and OPTIONS are supported."), 405);
                    return;
                }

                byte[] requestBytes = await ReadRequestBytesAsync(request.InputStream);
                string requestBody = DecodeRequestBody(request, requestBytes);

                JToken parsed;
                try
                {
                    parsed = ParseJsonToken(requestBody);
                }
                catch (Exception ex)
                {
                    RevitMcpDiagnosticLogger.LogException("POST JSON parse failed", ex);
                    await WriteJsonAsync(response, CreateError(null, -32700, "Invalid JSON: " + ex.Message), 400);
                    return;
                }

                JToken result;
                JArray batch = parsed as JArray;
                JObject single = parsed as JObject;
                if (batch != null)
                    result = await HandleBatchAsync(batch, request, cancellationToken);
                else if (single != null)
                    result = await HandleSingleAsync(single, request, cancellationToken);
                else
                    result = CreateError(null, -32600, "JSON-RPC request must be an object or an array.");

                await WriteJsonAsync(response, result, 200);
            }
            catch (Exception ex)
            {
                RevitMcpDiagnosticLogger.LogException("HandleContextAsync exception", ex);
                await TryWriteJsonAsync(response, CreateError(null, -32603, ex.Message), 500);
            }
            finally
            {
                try
                {
                    response.Close();
                }
                catch (Exception ex)
                {
                    RevitMcpDiagnosticLogger.LogException("Response close failed", ex);
                }
            }
        }

        private async Task<JArray> HandleBatchAsync(JArray batch, HttpListenerRequest httpRequest, CancellationToken cancellationToken)
        {
            JArray responses = new JArray();

            foreach (JToken item in batch)
            {
                JObject request = item as JObject;
                if (request == null)
                    responses.Add(CreateError(null, -32600, "Each JSON-RPC batch item must be an object."));
                else
                    responses.Add(await HandleSingleAsync(request, httpRequest, cancellationToken));
            }

            return responses;
        }

        private static JToken ParseJsonToken(string json)
        {
            using (StringReader stringReader = new StringReader(json ?? string.Empty))
            using (JsonTextReader jsonReader = new JsonTextReader(stringReader))
            {
                jsonReader.DateParseHandling = DateParseHandling.None;
                return JToken.ReadFrom(jsonReader);
            }
        }

        private async Task<JObject> HandleSingleAsync(JObject request, HttpListenerRequest httpRequest, CancellationToken cancellationToken)
        {
            JToken id = request["id"] == null ? JValue.CreateNull() : request["id"].DeepClone();
            string method = request["method"] == null ? null : request["method"].ToString();
            JObject parameters = request["params"] as JObject ?? new JObject();
            string scenario = ResolveInteractionScenario(httpRequest);

            switch (method)
            {
                case "initialize":
                    return CreateResult(id, HandleInitialize(parameters));

                case "ping":
                    return CreateResult(id, new JObject());

                case "tools/list":
                    return CreateResult(id, new JObject
                    {
                        { "tools", RevitMcpToolRegistry.ToMcpToolsArray() }
                    });

                case "tools/call":
                    return CreateResult(id, await HandleToolCallAsync(id, scenario, parameters, cancellationToken));

                case "notifications/initialized":
                    return CreateResult(id, new JObject());

                default:
                    RevitMcpDiagnosticLogger.Log("[MCP_TRANSPORT] unsupported method. Method=" + (method ?? "<null>") + ", scenario=" + scenario);
                    return CreateError(id, -32601, "MCP method is not supported: " + method);
            }
        }

        private JObject HandleInitialize(JObject parameters)
        {
            string protocolVersion = parameters["protocolVersion"] == null
                ? DefaultProtocolVersion
                : parameters["protocolVersion"].ToString();

            return new JObject
            {
                { "protocolVersion", protocolVersion },
                { "capabilities", new JObject { { "tools", new JObject() } } },
                { "serverInfo", new JObject { { "name", "revit-mcp" }, { "version", "0.1.0" } } }
            };
        }

        private async Task<JObject> HandleToolCallAsync(
            JToken id,
            string scenario,
            JObject parameters,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();

            string toolName = parameters["name"] == null ? null : parameters["name"].ToString();
            JObject arguments = parameters["arguments"] as JObject ?? new JObject();
            string requestId = CreateDiagnosticRequestId(id);
            string toolArea = GetToolArea(toolName);
            bool isInternalWpf = IsWpfScenario(scenario);
            RevitMcpInteractionLogger interactionLogger = isInternalWpf ? null : new RevitMcpInteractionLogger(scenario);
            Stopwatch stopwatch = Stopwatch.StartNew();

            RevitMcpDiagnosticLogger.Log("tools/call begin. ToolName=" + (toolName ?? "<null>") + ", scenario=" + scenario + ", requestId=" + requestId);
            if (interactionLogger != null)
                interactionLogger.LogToolStart(requestId, toolName, arguments, toolArea);

            RevitMcpToolCallResponse toolResponse;
            try
            {
                toolResponse = await _externalEventHandler.ExecuteToolAsync(
                    null,
                    null,
                    toolName,
                    arguments);
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                RevitMcpDiagnosticLogger.LogException("tools/call error. ToolName=" + (toolName ?? "<null>") + ", scenario=" + scenario + ", requestId=" + requestId, ex);
                if (interactionLogger != null)
                    interactionLogger.LogToolError(requestId, toolName, arguments, ex, stopwatch.ElapsedMilliseconds, toolArea);
                throw;
            }

            stopwatch.Stop();
            RevitMcpDiagnosticLogger.Log("tools/call end. ToolName=" + (toolName ?? "<null>") + ", Success=" + toolResponse.Success + ", scenario=" + scenario + ", requestId=" + requestId);
            if (!toolResponse.Success)
                RevitMcpDiagnosticLogger.Log("tools/call error result. ToolName=" + (toolName ?? "<null>") + ", scenario=" + scenario + ", requestId=" + requestId + ", ErrorCode=" + (toolResponse.Error == null ? string.Empty : toolResponse.Error.Code) + ", Message=" + (toolResponse.Error == null ? string.Empty : toolResponse.Error.Message));
            if (interactionLogger != null)
                interactionLogger.LogToolEnd(requestId, toolName, arguments, toolResponse, stopwatch.ElapsedMilliseconds, toolArea);

            string text = toolResponse.Success
                ? SerializeJsonToken(toolResponse.Result)
                : CreateToolErrorText(toolResponse);

            JObject result = new JObject
            {
                {
                    "content",
                    new JArray
                    {
                        new JObject
                        {
                            { "type", "text" },
                            { "text", text }
                        }
                    }
                },
                { "isError", !toolResponse.Success }
            };

            if (toolResponse.Success && toolResponse.Result != null)
                result["structuredContent"] = toolResponse.Result.DeepClone();
            else if (!toolResponse.Success)
                result["structuredContent"] = CreateToolErrorObject(toolResponse);
            return result;
        }

        private JObject CreateStatusObject()
        {
            return new JObject
            {
                { "name", "revit-mcp" },
                { "status", IsRunning ? "running" : "stopped" },
                { "endpoint", _prefix },
                { "tools", RevitMcpToolRegistry.GetAll().Count }
            };
        }

        private static JObject CreateResult(JToken id, JObject result)
        {
            return new JObject
            {
                { "jsonrpc", "2.0" },
                { "id", id == null ? JValue.CreateNull() : id },
                { "result", result ?? new JObject() }
            };
        }

        private static JObject CreateError(JToken id, int code, string message)
        {
            return new JObject
            {
                { "jsonrpc", "2.0" },
                { "id", id == null ? JValue.CreateNull() : id },
                { "error", new JObject { { "code", code }, { "message", message } } }
            };
        }

        private static async Task WriteJsonAsync(HttpListenerResponse response, JToken payload, int statusCode)
        {
            string json = SerializeJsonToken(payload);
            byte[] buffer = Encoding.UTF8.GetBytes(json);

            response.StatusCode = statusCode;
            response.ContentType = "application/json; charset=utf-8";
            response.ContentEncoding = Encoding.UTF8;
            response.ContentLength64 = buffer.Length;
            response.KeepAlive = false;

            await response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
            response.OutputStream.Flush();
        }

        private static async Task TryWriteJsonAsync(HttpListenerResponse response, JToken payload, int statusCode)
        {
            try
            {
                await WriteJsonAsync(response, payload, statusCode);
            }
            catch (Exception ex)
            {
                RevitMcpDiagnosticLogger.LogException("TryWriteJsonAsync failed", ex);
            }
        }

        private static void AddCorsHeaders(HttpListenerResponse response)
        {
            try
            {
                response.Headers["Access-Control-Allow-Origin"] = "http://127.0.0.1";
                response.Headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS";
                response.Headers["Access-Control-Allow-Headers"] = "Content-Type, Accept";
            }
            catch (Exception ex)
            {
                RevitMcpDiagnosticLogger.LogException("AddCorsHeaders failed", ex);
            }
        }

        private static Encoding ResolveRequestEncoding(HttpListenerRequest request)
        {
            if (request == null)
                return Encoding.UTF8;

            string contentType = request.ContentType ?? string.Empty;
            if (contentType.IndexOf("charset", StringComparison.OrdinalIgnoreCase) < 0)
                return Encoding.UTF8;

            try
            {
                return request.ContentEncoding ?? Encoding.UTF8;
            }
            catch
            {
                return Encoding.UTF8;
            }
        }

        private static async Task<byte[]> ReadRequestBytesAsync(Stream inputStream)
        {
            if (inputStream == null)
                return new byte[0];

            using (MemoryStream memoryStream = new MemoryStream())
            {
                await inputStream.CopyToAsync(memoryStream);
                return memoryStream.ToArray();
            }
        }

        private static string DecodeRequestBody(HttpListenerRequest request, byte[] requestBytes)
        {
            if (requestBytes == null || requestBytes.Length == 0)
            {
                return string.Empty;
            }

            Encoding requestEncoding = ResolveRequestEncoding(request);
            string decoded = requestEncoding.GetString(requestBytes);

            if (!HasExplicitCharset(request))
            {
                Encoding fallbackEncoding = ResolveBestFallbackEncoding(requestBytes, decoded);
                if (fallbackEncoding != null)
                {
                    string fallbackDecoded = fallbackEncoding.GetString(requestBytes);
                    return fallbackDecoded;
                }
            }

            return decoded;
        }

        private static bool HasExplicitCharset(HttpListenerRequest request)
        {
            if (request == null)
                return false;

            string contentType = request.ContentType ?? string.Empty;
            return contentType.IndexOf("charset", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static Encoding ResolveBestFallbackEncoding(byte[] requestBytes, string utf8Decoded)
        {
            Encoding bestEncoding = null;
            string bestDecoded = utf8Decoded;

            TrySelectFallbackEncoding(requestBytes, utf8Decoded, Encoding.Default, ref bestEncoding, ref bestDecoded);
            TrySelectFallbackEncoding(requestBytes, utf8Decoded, GetEncodingOrNull(1251), ref bestEncoding, ref bestDecoded);
            TrySelectFallbackEncoding(requestBytes, utf8Decoded, Encoding.Unicode, ref bestEncoding, ref bestDecoded);

            return bestEncoding;
        }

        private static void TrySelectFallbackEncoding(
            byte[] requestBytes,
            string utf8Decoded,
            Encoding candidateEncoding,
            ref Encoding bestEncoding,
            ref string bestDecoded)
        {
            if (requestBytes == null || requestBytes.Length == 0 || candidateEncoding == null)
                return;

            string candidateDecoded = candidateEncoding.GetString(requestBytes);
            if (ShouldUseFallbackEncoding(utf8Decoded, bestDecoded, candidateDecoded))
            {
                bestEncoding = candidateEncoding;
                bestDecoded = candidateDecoded;
            }
        }

        private static bool ShouldUseFallbackEncoding(string utf8Decoded, string currentBestDecoded, string fallbackDecoded)
        {
            if (string.IsNullOrEmpty(fallbackDecoded))
                return false;

            int currentReplacementChars = CountChar(currentBestDecoded, '\uFFFD');
            int fallbackReplacementChars = CountChar(fallbackDecoded, '\uFFFD');
            if (currentReplacementChars > fallbackReplacementChars)
                return true;

            int utf8CyrillicChars = CountCyrillicChars(utf8Decoded);
            int currentCyrillicChars = CountCyrillicChars(currentBestDecoded);
            int fallbackCyrillicChars = CountCyrillicChars(fallbackDecoded);
            return fallbackCyrillicChars > utf8CyrillicChars && fallbackCyrillicChars > currentCyrillicChars;
        }

        private static Encoding GetEncodingOrNull(int codePage)
        {
            try
            {
                return Encoding.GetEncoding(codePage);
            }
            catch (Exception ex)
            {
                RevitMcpDiagnosticLogger.LogException("Encoding.GetEncoding failed. CodePage=" + codePage, ex);
                return null;
            }
        }

        private static int CountChar(string value, char charToCount)
        {
            if (string.IsNullOrEmpty(value))
                return 0;

            int count = 0;
            foreach (char c in value)
            {
                if (c == charToCount)
                    count++;
            }

            return count;
        }

        private static int CountCyrillicChars(string value)
        {
            if (string.IsNullOrEmpty(value))
                return 0;

            int count = 0;
            foreach (char c in value)
            {
                if (c >= '\u0400' && c <= '\u04FF')
                    count++;
            }

            return count;
        }

        private static string SerializeJsonToken(JToken token)
        {
            if (token == null)
                return string.Empty;

            return JsonConvert.SerializeObject(token);
        }

        private static string CreateToolErrorText(RevitMcpToolCallResponse response)
        {
            if (response == null || response.Error == null)
                return "MCP tool execution failed.";

            string code = string.IsNullOrWhiteSpace(response.Error.Code)
                ? "tool_execution_failed"
                : response.Error.Code;

            string message = string.IsNullOrWhiteSpace(response.Error.Message)
                ? "MCP tool execution failed."
                : response.Error.Message;

            return code + ": " + message;
        }

        private static JObject CreateToolErrorObject(RevitMcpToolCallResponse response)
        {
            RevitMcpToolExecutionError error = response == null ? null : response.Error;

            JObject errorObject = new JObject
            {
                { "success", false },
                { "toolName", response == null ? string.Empty : (response.ToolName ?? string.Empty) },
                {
                    "error",
                    new JObject
                    {
                        { "code", error == null || string.IsNullOrWhiteSpace(error.Code) ? "tool_execution_failed" : error.Code },
                        { "message", error == null ? "MCP tool execution failed." : (error.Message ?? string.Empty) },
                        { "exceptionType", error == null ? string.Empty : (error.ExceptionType ?? string.Empty) },
                        { "details", error == null || error.Details == null ? new JObject() : (JObject)error.Details.DeepClone() }
                    }
                }
            };

            return errorObject;
        }

        private static string ResolveInteractionScenario(HttpListenerRequest request)
        {
            try
            {
                string clientName = request == null ? null : request.Headers["X-Revit-MCP-Client"];
                if (!string.IsNullOrWhiteSpace(clientName))
                    return NormalizeScenario(clientName);
            }
            catch
            {
            }

            return "external_mcp";
        }

        private static string NormalizeScenario(string value)
        {
            value = (value ?? string.Empty).Trim().ToLowerInvariant();
            if (string.IsNullOrWhiteSpace(value))
                return "external_mcp";

            StringBuilder builder = new StringBuilder();
            foreach (char c in value)
            {
                if ((c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '_' || c == '-')
                    builder.Append(c == '-' ? '_' : c);
            }

            return builder.Length == 0 ? "external_mcp" : builder.ToString();
        }

        private static bool IsWpfScenario(string scenario)
        {
            return string.Equals(scenario, "wpf_window", StringComparison.OrdinalIgnoreCase);
        }

        private static string CreateDiagnosticRequestId(JToken id)
        {
            string idText = id == null || id.Type == JTokenType.Null
                ? Guid.NewGuid().ToString("N")
                : id.ToString();

            return "mcp-" + NormalizeScenario(idText);
        }

        private static string GetToolArea(string toolName)
        {
            RevitMcpToolDefinition definition;
            if (!string.IsNullOrWhiteSpace(toolName) && RevitMcpToolRegistry.TryGet(toolName, out definition))
                return definition.Area;

            return string.Empty;
        }
    }
}

