using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace RevitMcpStdioProxy
{
    internal static class Program
    {
        private const string DefaultEndpoint = "http://127.0.0.1:48731/mcp/";
        private static readonly JavaScriptSerializer Serializer = new JavaScriptSerializer();

        private static async Task<int> Main(string[] args)
        {
            Console.InputEncoding = Encoding.UTF8;
            Console.OutputEncoding = new UTF8Encoding(false);

            string endpoint = ResolveEndpoint(args);
            Console.Error.WriteLine("Revit MCP stdio proxy started. Endpoint=" + endpoint);

            using (HttpClient client = new HttpClient())
            {
                client.Timeout = TimeSpan.FromMinutes(10);

                string line;
                while ((line = Console.ReadLine()) != null)
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    RequestMetadata metadata = ReadRequestMetadata(line);

                    try
                    {
                        string responseJson = await ForwardToRevitAsync(client, endpoint, line);
                        if (metadata.HasId)
                        {
                            Console.WriteLine(responseJson);
                            Console.Out.Flush();
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine("Proxy request failed: " + ex.GetType().Name + ": " + ex.Message);
                        if (metadata.HasId)
                        {
                            Console.WriteLine(CreateJsonRpcError(metadata.IdJson, -32000, "Revit MCP proxy error: " + ex.Message));
                            Console.Out.Flush();
                        }
                    }
                }
            }

            Console.Error.WriteLine("Revit MCP stdio proxy stopped.");
            return 0;
        }

        private static string ResolveEndpoint(string[] args)
        {
            for (int i = 0; args != null && i < args.Length - 1; i++)
            {
                if (string.Equals(args[i], "--endpoint", StringComparison.OrdinalIgnoreCase))
                    return NormalizeEndpoint(args[i + 1]);
            }

            string envEndpoint = Environment.GetEnvironmentVariable("REVIT_MCP_ENDPOINT");
            if (!string.IsNullOrWhiteSpace(envEndpoint))
                return NormalizeEndpoint(envEndpoint);

            return DefaultEndpoint;
        }

        private static string NormalizeEndpoint(string endpoint)
        {
            if (string.IsNullOrWhiteSpace(endpoint))
                return DefaultEndpoint;

            endpoint = endpoint.Trim();
            return endpoint.EndsWith("/", StringComparison.Ordinal) ? endpoint : endpoint + "/";
        }

        private static async Task<string> ForwardToRevitAsync(HttpClient client, string endpoint, string json)
        {
            using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, endpoint))
            {
                request.Content = new StringContent(json, Encoding.UTF8, "application/json");
                request.Headers.Add("X-Revit-MCP-Client", ResolveClientName());

                using (HttpResponseMessage response = await client.SendAsync(request))
                {
                    string responseJson = await response.Content.ReadAsStringAsync();
                    if (!response.IsSuccessStatusCode)
                        throw new InvalidOperationException("HTTP " + (int)response.StatusCode + ": " + responseJson);

                    return responseJson;
                }
            }
        }

        private static string ResolveClientName()
        {
            string value = Environment.GetEnvironmentVariable("REVIT_MCP_CLIENT_NAME");
            return string.IsNullOrWhiteSpace(value) ? "external_mcp" : value.Trim();
        }

        private static RequestMetadata ReadRequestMetadata(string json)
        {
            try
            {
                object parsed = Serializer.DeserializeObject(json);
                Dictionary<string, object> obj = parsed as Dictionary<string, object>;
                if (obj == null || !obj.ContainsKey("id"))
                    return new RequestMetadata(false, "null");

                return new RequestMetadata(true, Serializer.Serialize(obj["id"]));
            }
            catch
            {
                return new RequestMetadata(true, "null");
            }
        }

        private static string CreateJsonRpcError(string idJson, int code, string message)
        {
            if (string.IsNullOrWhiteSpace(idJson))
                idJson = "null";

            return "{\"jsonrpc\":\"2.0\",\"id\":"
                + idJson
                + ",\"error\":{\"code\":"
                + code
                + ",\"message\":"
                + Serializer.Serialize(message)
                + "}}";
        }

        private sealed class RequestMetadata
        {
            public RequestMetadata(bool hasId, string idJson)
            {
                HasId = hasId;
                IdJson = idJson;
            }

            public bool HasId { get; private set; }
            public string IdJson { get; private set; }
        }
    }
}
