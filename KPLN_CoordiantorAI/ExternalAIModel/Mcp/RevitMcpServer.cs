using System;
using System.Net;

namespace KPLN_CoordiantorAI.ExternalAIModel.Mcp
{
    internal sealed class RevitMcpServer : IDisposable
    {
        public const string DefaultEndpoint = "http://127.0.0.1:48731/mcp/";
        private const int FirstPort = 48731;
        private const int LastPort = 48799;

        private readonly RevitMcpExternalEventHandler _externalEventHandler;
        private readonly int _revitVersion;
        private RevitMcpTransport _transport;
        private bool _registered;
        private bool _disposed;

        public RevitMcpServer()
            : this(ModuleData.RevitVersion)
        {
        }

        public RevitMcpServer(int revitVersion)
        {
            _revitVersion = revitVersion;
            InstanceId = Guid.NewGuid().ToString("N");
            _externalEventHandler = new RevitMcpExternalEventHandler();
        }

        public string Endpoint { get; private set; }

        public string InstanceId { get; private set; }

        public bool IsRunning
        {
            get { return _transport != null && _transport.IsRunning; }
        }

        public void Start()
        {
            RevitMcpDiagnosticLogger.Log("RevitMcpServer.Start requested.");

            if (_disposed)
                throw new ObjectDisposedException(GetType().FullName);

            if (IsRunning)
                return;

            HttpListenerException lastPortException = null;
            for (int port = FirstPort; port <= LastPort; port++)
            {
                string endpoint = "http://127.0.0.1:" + port + "/mcp/";
                RevitMcpTransport candidate = new RevitMcpTransport(
                    endpoint,
                    _externalEventHandler,
                    InstanceId,
                    _revitVersion);

                try
                {
                    candidate.Start();
                }
                catch (HttpListenerException ex)
                {
                    if (!RevitMcpTransport.IsEndpointConflict(ex))
                        throw;

                    lastPortException = ex;
                    continue;
                }

                try
                {
                    RevitMcpInstanceRegistry.Register(InstanceId, endpoint, _revitVersion);
                    _transport = candidate;
                    Endpoint = endpoint;
                    _registered = true;
                    RevitMcpDiagnosticLogger.Log(
                        "RevitMcpServer.Start completed. IsRunning=True, Endpoint="
                        + Endpoint
                        + ", InstanceId="
                        + InstanceId
                        + ", RevitVersion="
                        + _revitVersion);
                    return;
                }
                catch
                {
                    candidate.Stop();
                    throw;
                }
            }

            throw new InvalidOperationException(
                "No free local MCP port was found in range "
                + FirstPort
                + "-"
                + LastPort
                + ".",
                lastPortException);
        }

        public void Stop()
        {
            RevitMcpDiagnosticLogger.Log("RevitMcpServer.Stop requested.");

            if (_disposed)
                return;

            if (_registered)
            {
                try
                {
                    RevitMcpInstanceRegistry.Unregister(InstanceId);
                }
                catch (Exception ex)
                {
                    RevitMcpDiagnosticLogger.LogException(
                        "Failed to unregister MCP instance. InstanceId=" + InstanceId,
                        ex);
                }

                _registered = false;
            }

            if (_transport != null)
            {
                _transport.Stop();
                _transport = null;
            }

            RevitMcpDiagnosticLogger.Log(
                "RevitMcpServer.Stop completed. InstanceId=" + InstanceId);
        }

        public void Dispose()
        {
            if (_disposed)
                return;

            Stop();
            _disposed = true;
            RevitMcpDiagnosticLogger.Log("RevitMcpServer disposed.");
        }
    }
}
