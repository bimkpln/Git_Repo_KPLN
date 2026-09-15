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
        private static readonly JavaScriptSerializer Serializer = new JavaScriptSerializer();

        private static async Task<int> Main(string[] args)
        {
            Console.InputEncoding = Encoding.UTF8;
            Console.OutputEncoding = new UTF8Encoding(false);

            ProxyOptions options;
            try
            {
                options = ProxyOptions.Parse(args);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Invalid proxy configuration: " + ex.Message);
                return 2;
            }

            Console.Error.WriteLine("Revit MCP stdio proxy started. Mode=" + options.Describe());

            using (HttpClient client = new HttpClient())
            {
                client.Timeout = TimeSpan.FromMinutes(10);
                RevitMcpRouter router = new RevitMcpRouter(client, options);

                string line;
                while ((line = Console.ReadLine()) != null)
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    RequestMetadata metadata = ReadRequestMetadata(line);

                    try
                    {
                        string responseJson = await router.HandleAsync(line);
                        if (metadata.HasId && !string.IsNullOrWhiteSpace(responseJson))
                        {
                            Console.WriteLine(responseJson);
                            Console.Out.Flush();
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine(
                            "Proxy request failed: "
                            + ex.GetType().Name
                            + ": "
                            + ex.Message);
                        if (metadata.HasId)
                        {
                            Console.WriteLine(
                                CreateJsonRpcError(
                                    metadata.IdJson,
                                    -32000,
                                    "Revit MCP proxy error: " + ex.Message));
                            Console.Out.Flush();
                        }
                    }
                }
            }

            Console.Error.WriteLine("Revit MCP stdio proxy stopped.");
            return 0;
        }

        private static RequestMetadata ReadRequestMetadata(string json)
        {
            try
            {
                object parsed = Serializer.DeserializeObject(json);
                Dictionary<string, object> obj = parsed as Dictionary<string, object>;
                if (obj == null || !obj.ContainsKey("id"))
                    return new RequestMetadata(obj == null, "null");

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
