using System;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Threading;
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
            JObject arguments,
            CancellationToken cancellationToken)
        {
            if (cancellationToken.IsCancellationRequested)
                return CreateCanceledTask();

            PendingToolCall request = new PendingToolCall
            {
                Document = document,
                UiDocument = uiDocument,
                ToolName = toolName,
                Arguments = arguments == null ? new JObject() : (JObject)arguments.DeepClone(),
                Completion = new TaskCompletionSource<RevitMcpToolCallResponse>(
                    TaskCreationOptions.RunContinuationsAsynchronously),
                CancellationToken = cancellationToken
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

            request.CancellationRegistration = cancellationToken.Register(
                () => CancelRequest(request));
            request.Completion.Task.ContinueWith(
                task => request.CancellationRegistration.Dispose(),
                CancellationToken.None,
                TaskContinuationOptions.ExecuteSynchronously,
                TaskScheduler.Default);

            if (request.Completion.Task.IsCompleted)
                return request.Completion.Task;

            try
            {
                _externalEvent.Raise();
            }
            catch (Exception ex)
            {
                ClearPendingRequest(request);
                request.Completion.TrySetResult(new RevitMcpToolCallResponse
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
                if (request.CancellationToken.IsCancellationRequested)
                {
                    request.Completion.TrySetCanceled();
                    return;
                }

                UIDocument uiDocument = request.UiDocument ?? app.ActiveUIDocument;
                Document document = request.Document ?? (uiDocument == null ? null : uiDocument.Document);

                RevitMcpToolExecutor executor = new RevitMcpToolExecutor(
                    document,
                    uiDocument);

                RevitMcpToolCallResponse response = executor.Execute(
                    request.ToolName,
                    request.Arguments,
                    request.CancellationToken);
                FillRevitContext(response, document, uiDocument);
                if (request.CancellationToken.IsCancellationRequested)
                    request.Completion.TrySetCanceled();
                else
                    request.Completion.TrySetResult(response);
            }
            catch (Exception ex)
            {
                if (request.CancellationToken.IsCancellationRequested)
                {
                    request.Completion.TrySetCanceled();
                    return;
                }

                request.Completion.TrySetResult(new RevitMcpToolCallResponse
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

        private void CancelRequest(PendingToolCall request)
        {
            lock (_syncRoot)
            {
                if (ReferenceEquals(_pendingToolCall, request))
                    _pendingToolCall = null;
            }

            request.Completion.TrySetCanceled();
        }

        private static Task<RevitMcpToolCallResponse> CreateCanceledTask()
        {
            TaskCompletionSource<RevitMcpToolCallResponse> completion =
                new TaskCompletionSource<RevitMcpToolCallResponse>();
            completion.SetCanceled();
            return completion.Task;
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
            public CancellationToken CancellationToken { get; set; }
            public CancellationTokenRegistration CancellationRegistration { get; set; }
        }
    }

    internal static class RevitMcpRequestCancellationRegistry
    {
        private static readonly object SyncRoot = new object();
        private static readonly Dictionary<string, CancellationToken> Tokens =
            new Dictionary<string, CancellationToken>(StringComparer.Ordinal);

        public static void Register(string requestScopeId, CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(requestScopeId))
                return;

            lock (SyncRoot)
                Tokens[requestScopeId] = cancellationToken;
        }

        public static CancellationToken GetToken(string requestScopeId)
        {
            if (string.IsNullOrWhiteSpace(requestScopeId))
                return CancellationToken.None;

            lock (SyncRoot)
            {
                CancellationToken token;
                return Tokens.TryGetValue(requestScopeId, out token)
                    ? token
                    : CancellationToken.None;
            }
        }

        public static void Unregister(string requestScopeId)
        {
            if (string.IsNullOrWhiteSpace(requestScopeId))
                return;

            lock (SyncRoot)
                Tokens.Remove(requestScopeId);
        }
    }
}


