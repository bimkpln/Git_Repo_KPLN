using System.Diagnostics;
using System.IO.Pipes;
using System.Text.RegularExpressions;

namespace KPLN_WebApplication
{
    public class Program
    {
        public static void Main(string[] args)
        {
            if (int.TryParse(args[0], out int port))
            { 
                var builder = WebApplication.CreateBuilder(args);
                var app = builder.Build();

                app.MapGet("/select/{elementIds}", async (HttpContext context, string elementIds) =>
                {
                    List<int> ids = Regex
                        .Split(elementIds, "[^0-9]+")
                        .Where(s => !string.IsNullOrWhiteSpace(s))
                        .Select(int.Parse)
                        .ToList();

                    bool success = await TrySendElementIdsToAllRevitInstances(ids);
                    if (success)
                    {
                        return Results.Ok($"Элементы с Id: {string.Join(", ", ids)} успешно выделены в модели. Перейди в Revit.");
                    }
                    return Results.Problem("Ошибка отправки.");
                });

                app.Run($"http://localhost:{port}");
            }
        }

        private static async Task<bool> TrySendElementIdsToAllRevitInstances(List<int> elementIds)
        {
            bool atLeastOneSuccess = false;

            foreach (var process in Process.GetProcessesByName("Revit"))
            {
                int revitProcessId = process.Id;
                string pipeName = $"RevitPipe_{revitProcessId}";

                if (await SendElementIdToRevit(elementIds, pipeName))
                    atLeastOneSuccess = true;
            }

            return atLeastOneSuccess;
        }

        private static async Task<bool> SendElementIdToRevit(List<int> elementIds, string pipeName)
        {
            try
            {
                using var pipeClient = new NamedPipeClientStream(".", pipeName, PipeDirection.Out);
                await pipeClient.ConnectAsync(1000); // Чакаем 1 сек

                using var writer = new StreamWriter(pipeClient);
                string elemIds = string.Join(",", elementIds);
                await writer.WriteLineAsync(elemIds);
                await writer.FlushAsync();

                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
