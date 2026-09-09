using System;
using System.Collections.Generic;
using KPLN_CoordiantorAI.ExternalAIModel;

namespace KPLN_CoordiantorAI.ExternalAIModel.Mcp
{
    internal static class RevitMcpDiagnosticLogger
    {
        private static readonly object SyncRoot = new object();
        private static DiagnosticLogger _logger;

        public static void Log(string message)
        {
            try
            {
                GetLogger().LogEvent("mcp", "[MCP] " + (message ?? string.Empty));
            }
            catch
            {
            }
        }

        public static void LogException(string area, Exception ex)
        {
            try
            {
                GetLogger().LogException(
                    "mcp",
                    "[MCP] " + (area ?? "EXCEPTION"),
                    ex,
                    new Dictionary<string, object>());
            }
            catch
            {
            }
        }

        private static DiagnosticLogger GetLogger()
        {
            lock (SyncRoot)
            {
                if (_logger == null)
                    _logger = new DiagnosticLogger(null);

                return _logger;
            }
        }
    }
}
