using System;
using System.Collections.Generic;
using System.Threading;
using KPLN_CoordiantorAI.Common;
using KPLN_CoordiantorAI.ExternalModel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace KPLN_CoordiantorAI.ExternalAIModel.Mcp
{
    internal sealed class RevitMcpInteractionLogger
    {
        private const int ChatLogFlushDelayMs = 2000;
        private static readonly object ChatBatchSyncRoot = new object();
        private static List<McpToolLogItem> ChatBatchItems = new List<McpToolLogItem>();
        private static Dictionary<string, int> ChatBatchToolAreaStats = new Dictionary<string, int>();
        private static DateTime ChatBatchRequestTime;
        private static DateTime ChatBatchResponseTime;
        private static string ChatBatchScenario;
        private static string ChatBatchModelName;
        private static string ChatBatchViewName;
        private static Timer ChatBatchFlushTimer;

        private readonly string _scenario;
        private readonly DiagnosticLogger _diagnosticLogger;

        public RevitMcpInteractionLogger(string scenario)
        {
            _scenario = string.IsNullOrWhiteSpace(scenario) ? "external_mcp" : scenario.Trim();
            _diagnosticLogger = new DiagnosticLogger(null, _scenario);
        }

        public void LogToolStart(string requestId, string toolName, JObject arguments, string toolArea)
        {
            _diagnosticLogger.LogEvent(requestId, "TOOL.START", new Dictionary<string, object>
            {
                { "toolName", toolName },
                { "argsLength", Serialize(arguments).Length },
                { "toolArea", toolArea },
                { "executionPath", "mcp" }
            });
        }

        public void LogToolEnd(
            string requestId,
            string toolName,
            JObject arguments,
            RevitMcpToolCallResponse response,
            long elapsedMs,
            string toolArea)
        {
            string resultText = response == null || response.Result == null
                ? string.Empty
                : JsonConvert.SerializeObject(response.Result);

            _diagnosticLogger.LogEvent(requestId, "TOOL.END", new Dictionary<string, object>
            {
                { "toolName", toolName },
                { "resultLength", resultText.Length },
                { "elapsedMs", elapsedMs },
                { "executionPath", "mcp" },
                { "success", response != null && response.Success }
            });

            DateTime responseTime = DateTime.Now;
            DateTime requestTime = responseTime.AddMilliseconds(-elapsedMs);
            EnqueueChatLogItem(
                new McpToolLogItem
                {
                    ToolName = toolName,
                    Success = response != null && response.Success,
                    ElapsedMs = elapsedMs,
                    ArgumentsLength = Serialize(arguments).Length,
                    ResultLength = resultText.Length,
                    ErrorMessage = GetErrorMessage(response)
                },
                requestTime,
                responseTime,
                response == null ? null : response.RevitModelName,
                response == null ? null : response.RevitViewName,
                toolArea);
        }

        public void LogToolError(
            string requestId,
            string toolName,
            JObject arguments,
            Exception exception,
            long elapsedMs,
            string toolArea)
        {
            _diagnosticLogger.LogException(requestId, "TOOL.ERROR", exception, new Dictionary<string, object>
            {
                { "toolName", toolName },
                { "elapsedMs", elapsedMs },
                { "executionPath", "mcp" }
            });

            DateTime responseTime = DateTime.Now;
            DateTime requestTime = responseTime.AddMilliseconds(-elapsedMs);
            EnqueueChatLogItem(
                new McpToolLogItem
                {
                    ToolName = toolName,
                    Success = false,
                    ElapsedMs = elapsedMs,
                    ArgumentsLength = Serialize(arguments).Length,
                    ResultLength = 0,
                    ErrorMessage = exception == null ? string.Empty : exception.Message
                },
                requestTime,
                responseTime,
                null,
                null,
                toolArea);
        }

        private static string LoadExternalModelLogFolder()
        {
            try
            {
                CoordinatorAiRepository repository = new CoordinatorAiRepository();
                ExternalModelSettings settings = repository.LoadExternalModelSettings();
                return settings == null ? string.Empty : settings.LogFolder;
            }
            catch (Exception ex)
            {
                RevitMcpDiagnosticLogger.LogException("LoadExternalModelLogFolder failed", ex);
                return string.Empty;
            }
        }

        private static IDictionary<string, int> BuildToolAreaStats(string toolArea)
        {
            Dictionary<string, int> stats = new Dictionary<string, int>();
            if (!string.IsNullOrWhiteSpace(toolArea))
                stats[toolArea] = 1;

            return stats;
        }

        private void EnqueueChatLogItem(
            McpToolLogItem item,
            DateTime requestTime,
            DateTime responseTime,
            string revitModelName,
            string revitViewName,
            string toolArea)
        {
            if (item == null)
                return;

            lock (ChatBatchSyncRoot)
            {
                if (ChatBatchItems.Count == 0)
                {
                    ChatBatchRequestTime = requestTime;
                    ChatBatchScenario = _scenario;
                    ChatBatchModelName = revitModelName;
                    ChatBatchViewName = revitViewName;
                    ChatBatchToolAreaStats = new Dictionary<string, int>();
                }

                ChatBatchItems.Add(item);
                ChatBatchResponseTime = responseTime;

                if (!string.IsNullOrWhiteSpace(revitModelName))
                    ChatBatchModelName = revitModelName;

                if (!string.IsNullOrWhiteSpace(revitViewName))
                    ChatBatchViewName = revitViewName;

                if (!string.IsNullOrWhiteSpace(toolArea))
                {
                    int count;
                    ChatBatchToolAreaStats.TryGetValue(toolArea, out count);
                    ChatBatchToolAreaStats[toolArea] = count + 1;
                }

                if (ChatBatchFlushTimer == null)
                    ChatBatchFlushTimer = new Timer(FlushChatBatch, null, ChatLogFlushDelayMs, Timeout.Infinite);
                else
                    ChatBatchFlushTimer.Change(ChatLogFlushDelayMs, Timeout.Infinite);
            }
        }

        private static void FlushChatBatch(object state)
        {
            List<McpToolLogItem> items;
            Dictionary<string, int> stats;
            DateTime requestTime;
            DateTime responseTime;
            string scenario;
            string modelName;
            string viewName;

            lock (ChatBatchSyncRoot)
            {
                if (ChatBatchItems.Count == 0)
                    return;

                items = new List<McpToolLogItem>(ChatBatchItems);
                stats = new Dictionary<string, int>(ChatBatchToolAreaStats);
                requestTime = ChatBatchRequestTime;
                responseTime = ChatBatchResponseTime;
                scenario = ChatBatchScenario;
                modelName = ChatBatchModelName;
                viewName = ChatBatchViewName;

                ChatBatchItems.Clear();
                ChatBatchToolAreaStats.Clear();
                ChatBatchScenario = null;
                ChatBatchModelName = null;
                ChatBatchViewName = null;
            }

            ChatLogger chatLogger = new ChatLogger(LoadExternalModelLogFolder(), scenario);
            chatLogger.LogMcpToolBatch(items, requestTime, responseTime, modelName, viewName, stats);
        }

        private static string Serialize(JToken token)
        {
            return token == null ? "{}" : JsonConvert.SerializeObject(token);
        }

        private static string GetErrorMessage(RevitMcpToolCallResponse response)
        {
            if (response == null || response.Success || response.Error == null)
                return string.Empty;

            return response.Error.Message;
        }
    }
}
