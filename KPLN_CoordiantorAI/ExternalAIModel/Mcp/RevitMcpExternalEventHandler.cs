using System;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json.Linq;

namespace KPLN_CoordiantorAI.ExternalAIModel.Mcp
{
    internal class RevitMcpExternalEventHandler : IExternalEventHandler
    {
        private readonly object _syncRoot = new object();
        private readonly ExternalEvent _externalEvent;
        private PendingToolCall _pendingToolCall;

        public RevitMcpExternalEventHandler()
        {
            _externalEvent = ExternalEvent.Create(this);
        }

        public Task<RevitMcpToolCallResponse> ExecuteToolAsync(
            Document document,
            UIDocument uiDocument,
            string toolName,
            JObject arguments)
        {
            PendingToolCall request = new PendingToolCall
            {
                Document = document,
                UiDocument = uiDocument,
                ToolName = toolName,
                Arguments = arguments == null ? new JObject() : (JObject)arguments.DeepClone(),
                Completion = new TaskCompletionSource<RevitMcpToolCallResponse>()
            };

            lock (_syncRoot)
            {
                if (_pendingToolCall != null)
                {
                    request.Completion.SetResult(new RevitMcpToolCallResponse
                    {
                        Success = false,
                        ToolName = toolName,
                        Error = new RevitMcpToolExecutionError
                        {
                            Code = "revit_busy",
                            Message = "Revit is already executing another MCP tool call. Retry in a few seconds.",
                            Details = new JObject { { "toolName", toolName ?? string.Empty } }
                        }
                    });

                    return request.Completion.Task;
                }

                _pendingToolCall = request;
            }

            try
            {
                _externalEvent.Raise();
            }
            catch (Exception ex)
            {
                ClearPendingRequest(request);
                request.Completion.SetResult(new RevitMcpToolCallResponse
                {
                    Success = false,
                    ToolName = toolName,
                    Error = new RevitMcpToolExecutionError
                    {
                        Code = "external_event_raise_failed",
                        Message = "Failed to pass MCP tool call into Revit API context: " + ex.Message,
                        ExceptionType = ex.GetType().FullName,
                        Details = new JObject { { "toolName", toolName ?? string.Empty } }
                    }
                });
            }

            return request.Completion.Task;
        }
        public void Execute(UIApplication app)
        {
            PendingToolCall request = TakePendingRequest();
            if (request == null)
                return;

            try
            {
                UIDocument uiDocument = request.UiDocument ?? app.ActiveUIDocument;
                Document document = request.Document ?? (uiDocument == null ? null : uiDocument.Document);

                RevitMcpToolExecutor executor = new RevitMcpToolExecutor(
                    document,
                    uiDocument);

                RevitMcpToolCallResponse response = executor.Execute(request.ToolName, request.Arguments);
                FillRevitContext(response, document, uiDocument);
                request.Completion.SetResult(response);
            }
            catch (Exception ex)
            {
                request.Completion.SetResult(new RevitMcpToolCallResponse
                {
                    Success = false,
                    ToolName = request.ToolName,
                    Error = new RevitMcpToolExecutionError
                    {
                        Code = "revit_api_context_error",
                        Message = "MCP tool failed inside Revit API context: " + ex.Message,
                        ExceptionType = ex.GetType().FullName,
                        Details = new JObject { { "toolName", request.ToolName ?? string.Empty } }
                    }
                });
            }
        }

        public string GetName()
        {
            return "Coordinator AI Revit MCP External Event Handler";
        }

        private PendingToolCall TakePendingRequest()
        {
            lock (_syncRoot)
            {
                PendingToolCall request = _pendingToolCall;
                _pendingToolCall = null;
                return request;
            }
        }

        private void ClearPendingRequest(PendingToolCall request)
        {
            lock (_syncRoot)
            {
                if (ReferenceEquals(_pendingToolCall, request))
                    _pendingToolCall = null;
            }
        }

        private static void FillRevitContext(RevitMcpToolCallResponse response, Document document, UIDocument uiDocument)
        {
            if (response == null)
                return;

            try
            {
                if (document != null)
                    response.RevitModelName = document.Title;

                if (uiDocument != null && uiDocument.ActiveView != null)
                    response.RevitViewName = uiDocument.ActiveView.Name;
            }
            catch
            {
            }
        }

        private sealed class PendingToolCall
        {
            public Document Document { get; set; }
            public UIDocument UiDocument { get; set; }
            public string ToolName { get; set; }
            public JObject Arguments { get; set; }
            public TaskCompletionSource<RevitMcpToolCallResponse> Completion { get; set; }
        }
    }
}


