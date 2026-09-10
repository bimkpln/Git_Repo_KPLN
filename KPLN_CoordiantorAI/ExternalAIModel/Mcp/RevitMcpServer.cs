using System;

namespace KPLN_CoordiantorAI.ExternalAIModel.Mcp
{
    internal sealed class RevitMcpServer : IDisposable
    {
        public const string DefaultEndpoint = "http://127.0.0.1:48731/mcp/";

        private readonly RevitMcpExternalEventHandler _externalEventHandler;
        private readonly RevitMcpTransport _transport;
        private bool _disposed;

        public RevitMcpServer()
        {
            _externalEventHandler = new RevitMcpExternalEventHandler();
            _transport = new RevitMcpTransport(DefaultEndpoint, _externalEventHandler);
        }

        public bool IsRunning
        {
            get { return _transport.IsRunning; }
        }

        public void Start()
        {
            RevitMcpDiagnosticLogger.Log("RevitMcpServer.Start requested.");

            if (_disposed)
                throw new ObjectDisposedException(GetType().FullName);

            _transport.Start();
            RevitMcpDiagnosticLogger.Log("RevitMcpServer.Start completed. IsRunning=" + IsRunning);
        }

        public void Stop()
        {
            RevitMcpDiagnosticLogger.Log("RevitMcpServer.Stop requested.");

            if (_disposed)
                return;

            _transport.Stop();
            RevitMcpDiagnosticLogger.Log("RevitMcpServer.Stop completed.");
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
