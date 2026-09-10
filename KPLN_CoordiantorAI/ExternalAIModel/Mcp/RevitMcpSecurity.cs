using System.Net;

namespace KPLN_CoordiantorAI.ExternalAIModel.Mcp
{
    internal static class RevitMcpSecurity
    {
        public static bool IsLocalRequest(HttpListenerRequest request)
        {
            if (request == null || request.RemoteEndPoint == null)
                return false;

            return IPAddress.IsLoopback(request.RemoteEndPoint.Address);
        }
    }
}
