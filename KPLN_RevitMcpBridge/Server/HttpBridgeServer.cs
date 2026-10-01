using Autodesk.Revit.UI;
using KPLN_RevitMcpBridge.Services;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace KPLN_RevitMcpBridge.Server
{
    internal sealed class HttpBridgeServer : IDisposable, IExternalEventHandler
    {
        private readonly UIControlledApplication _application;
        private readonly RevitService _service = new RevitService();
        private readonly ConcurrentQueue<WorkItem> _queue = new ConcurrentQueue<WorkItem>();
        private readonly Dictionary<string, WorkItem> _operations = new Dictionary<string, WorkItem>();
        private readonly object _gate = new object();
        private readonly SemaphoreSlim _clients = new SemaphoreSlim(16);
        private readonly CancellationTokenSource _shutdown = new CancellationTokenSource();
        private readonly string _token;
        private readonly string _version;
        private readonly string _sessionsDirectory;
        private readonly int _commandTimeoutMs;
        private readonly Action<IntPtr> _wakeWindow;
        private ExternalEvent _externalEvent;
        private Timer _dispatchTimer;
        private IntPtr _mainWindowHandle;
        private bool _executing;
        private HttpListener _listener;
        private string _discoveryPath;
        private volatile bool _stopped;
        public string SessionId { get; } = Guid.NewGuid().ToString("N");
        public string Url { get; private set; }

        public HttpBridgeServer(UIControlledApplication application, string sessionsDirectory = null, int commandTimeoutMs = 30000, Action<IntPtr> wakeWindow = null)
        {
            _application = application;
            _sessionsDirectory = sessionsDirectory;
            _commandTimeoutMs = commandTimeoutMs;
            _wakeWindow = wakeWindow ?? WakeWindow;
            _version = application.ControlledApplication.VersionNumber;
            var bytes = new byte[32];
            using (var random = RandomNumberGenerator.Create()) random.GetBytes(bytes);
            _token = Convert.ToBase64String(bytes);
        }

        public void Start()
        {
            // Start is called by the ribbon command in a valid Revit API context.
            // Cache the HWND here: HTTP/timer threads must not read Revit objects.
            _mainWindowHandle = _application.MainWindowHandle;
            _externalEvent = ExternalEvent.Create(this);
            _dispatchTimer = new Timer(RequestDispatch, null, Timeout.Infinite, Timeout.Infinite);
            Exception last = null;
            for (int port = 18765; port < 18785; port++)
            {
                var candidate = new HttpListener();
                candidate.Prefixes.Add("http://127.0.0.1:" + port + "/");
                try { candidate.Start(); _listener = candidate; Url = "http://127.0.0.1:" + port; break; }
                catch (HttpListenerException ex) { last = ex; candidate.Close(); }
            }
            if (_listener == null) throw new InvalidOperationException("Не удалось открыть localhost:18765–18784. " + last?.Message);
            _application.ControlledApplication.DocumentChanged += _service.DocumentChanged;
            var directory = _sessionsDirectory ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "KPLN", "RevitMcpBridge", "sessions");
            Directory.CreateDirectory(directory);
            _discoveryPath = Path.Combine(directory, SessionId + ".json");
            var temp = _discoveryPath + ".tmp";
            File.WriteAllText(temp, Json.Serialize(new { session_id = SessionId, url = Url, token = _token, pid = Process.GetCurrentProcess().Id, revit_version = _version, protocol = 1 }), new UTF8Encoding(false));
            File.Move(temp, _discoveryPath);
            _ = AcceptLoop();
        }

        private async Task AcceptLoop()
        {
            while (!_stopped)
            {
                try
                {
                    var context = await _listener.GetContextAsync().ConfigureAwait(false);
                    if (!_clients.Wait(0)) { context.Response.StatusCode = 503; context.Response.Close(); continue; }
                    _ = Handle(context);
                }
                catch (Exception) when (_stopped) { break; }
                catch (Exception ex) { Trace.TraceError("Revit MCP listener: " + ex.Message); break; }
            }
        }

        private bool Authorized(string value)
        {
            var expected = "Bearer " + _token;
            if (value == null || value.Length != expected.Length) return false;
            int diff = 0;
            for (int i = 0; i < value.Length; i++) diff |= value[i] ^ expected[i];
            return diff == 0;
        }

        private async Task Handle(HttpListenerContext context)
        {
            try
            {
                var request = context.Request;
                if (!request.IsLocal || request.Headers["Origin"] != null || !Authorized(request.Headers["Authorization"]))
                    throw new BridgeException("unauthorized", "Нет доступа к локальной сессии.", 401);
                object result;
                if (request.HttpMethod == "GET" && request.Url.AbsolutePath == "/health")
                    result = new { ok = true, session_id = SessionId, revit_version = _version, protocol = 1, bridge_version = "1.0.7", api_context = "ExternalEvent", background_wakeup = "WM_NULL" };
                else if (request.HttpMethod == "GET" && request.Url.AbsolutePath.StartsWith("/operations/", StringComparison.Ordinal))
                {
                    WorkItem item;
                    lock (_gate) _operations.TryGetValue(request.Url.AbsolutePath.Substring(12), out item);
                    if (item == null) throw new BridgeException("operation_not_found", "Запрос неизвестен или срок хранения результата истёк.", 404);
                    result = new { ok = true, operation_id = item.Id, state = item.State, response = item.Completion.Task.IsCompleted ? item.Completion.Task.Result : null };
                }
                else if (request.HttpMethod == "POST" && request.Url.AbsolutePath == "/command")
                {
                    if (request.ContentLength64 < 1 || request.ContentLength64 > 1024 * 1024)
                        throw new BridgeException("body_size", "Требуется Content-Length от 1 до 1048576 байт.", 413);
                    if (!(request.ContentType ?? "").StartsWith("application/json", StringComparison.OrdinalIgnoreCase))
                        throw new BridgeException("content_type", "Требуется application/json.", 415);
                    string body;
                    using (var reader = new StreamReader(request.InputStream, Encoding.UTF8))
                    {
                        var read = reader.ReadToEndAsync();
                        if (await Task.WhenAny(read, Task.Delay(5000, _shutdown.Token)).ConfigureAwait(false) != read)
                            throw new BridgeException("read_timeout", "Таймаут чтения запроса.", 408);
                        body = await read.ConfigureAwait(false);
                    }
                    var input = Json.Parse(body);
                    Guid operationGuid;
                    if (!Guid.TryParse(input.Text("operation_id", true), out operationGuid)) throw new BridgeException("invalid_input", "operation_id должен быть UUID.");
                    if (input.Text("session_id", true) != SessionId) throw new BridgeException("wrong_session", "Сессия Revit сменилась.", 409);
                    string id = operationGuid.ToString("N");
                    WorkItem item;
                    lock (_gate)
                    {
                        if (_stopped) throw new BridgeException("stopped", "Мост остановлен.", 503);
                        if (_operations.TryGetValue(id, out item))
                        {
                            if (item.Payload != body) throw new BridgeException("operation_conflict", "UUID уже использован для другого запроса.", 409);
                        }
                        else
                        {
                            foreach (var key in _operations.Where(p => p.Value.Completion.Task.IsCompleted && p.Value.Created < DateTime.UtcNow.AddMinutes(-15)).Select(p => p.Key).ToArray()) _operations.Remove(key);
                            if (_operations.Count >= 512 || _queue.Count >= 32) throw new BridgeException("busy", "Очередь заполнена; дождитесь завершения запросов.", 503);
                            item = new WorkItem(id, body); _operations.Add(id, item); _queue.Enqueue(item);
                        }
                    }
                    RequestDispatch(null);
                    if (await Task.WhenAny(item.Completion.Task, Task.Delay(_commandTimeoutMs, _shutdown.Token)).ConfigureAwait(false) == item.Completion.Task)
                        result = await item.Completion.Task.ConfigureAwait(false);
                    else
                    {
                        bool cancelled = item.Cancel();
                        result = new { ok = false, operation_id = id, error = new { code = cancelled ? "queue_timeout" : "outcome_pending", message = cancelled ? "Revit занят. Запрос отменён до выполнения." : "Команда уже исполняется. Прочитайте get_operation; не повторяйте запись." } };
                    }
                }
                else throw new BridgeException("not_found", "Неизвестный маршрут.", 404);
                await Write(context, 200, result).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                try { await Write(context, (ex as BridgeException)?.Status ?? 500, Json.Error(ex)).ConfigureAwait(false); }
                catch (Exception io) { Trace.TraceWarning("Revit MCP response: " + io.Message); }
            }
            finally { context.Response.Close(); _clients.Release(); }
        }

        private static async Task Write(HttpListenerContext context, int status, object value)
        {
            var bytes = Encoding.UTF8.GetBytes(Json.Serialize(value));
            context.Response.StatusCode = status;
            context.Response.ContentType = "application/json; charset=utf-8";
            context.Response.ContentLength64 = bytes.Length;
            await context.Response.OutputStream.WriteAsync(bytes, 0, bytes.Length).ConfigureAwait(false);
        }

        [DllImport("user32.dll", EntryPoint = "PostMessageW", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool PostMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);

        private static void WakeWindow(IntPtr window)
        {
            // WM_NULL only wakes the message loop; it never activates the window
            // or bypasses a modal dialog / an active Revit command.
            if (window != IntPtr.Zero && !PostMessage(window, 0, IntPtr.Zero, IntPtr.Zero))
                Trace.TraceWarning("Revit MCP wakeup: " + Marshal.GetLastWin32Error());
        }

        private void RequestDispatch(object state)
        {
            lock (_gate)
            {
                if (_stopped) return;
                WorkItem cancelled;
                while (_queue.TryPeek(out cancelled) && cancelled.Completion.Task.IsCompleted)
                    _queue.TryDequeue(out cancelled);
                if (_queue.IsEmpty) return;
                try
                {
                    if (!_executing)
                    {
                        // Raise is the only Revit API entry allowed from a worker
                        // thread. Actual model access stays in Execute below.
                        _externalEvent.Raise();
                        _wakeWindow(_mainWindowHandle);
                    }
                }
                catch (Exception ex) { Trace.TraceWarning("Revit MCP dispatch: " + ex.Message); }
                // Pending/TimedOut and requests arriving while the event returns
                // need another signal. Retry only while queued work exists.
                _dispatchTimer.Change(250, Timeout.Infinite);
            }
        }

        public string GetName() => "KPLN Revit MCP request queue";

        public void Execute(UIApplication app)
        {
            WorkItem item;
            lock (_gate)
            {
                if (_stopped || !_queue.TryDequeue(out item)) return;
                _executing = true;
            }
            try { item.Execute(() => _service.Execute(app, Json.Parse(item.Payload))); }
            finally
            {
                lock (_gate)
                {
                    _executing = false;
                    // Yield to Revit between commands. Do not re-raise the same
                    // event inside its handler; the timer retries after return.
                    if (!_stopped && !_queue.IsEmpty) _dispatchTimer.Change(1, Timeout.Infinite);
                }
            }
        }

        public void Dispose()
        {
            lock (_gate)
            {
                if (_stopped) return;
                _stopped = true;
                _shutdown.Cancel();
                foreach (var item in _operations.Values) item.Cancel();
                _dispatchTimer?.Dispose();
                _externalEvent?.Dispose();
            }
            _application.ControlledApplication.DocumentChanged -= _service.DocumentChanged;
            _listener?.Close();
            if (_discoveryPath != null)
                try { File.Delete(_discoveryPath); } catch (IOException ex) { Trace.TraceWarning(ex.Message); }
        }
    }
}
