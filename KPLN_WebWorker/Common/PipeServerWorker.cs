using Autodesk.Revit.DB;
using KPLN_WebWorker.ExecutableCommand;
using System;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Threading;

namespace KPLN_WebWorker.Common
{
    public static class PipeServerWorker
    {
        public static void StartListening()
        {
            int processId = Process.GetCurrentProcess().Id;
            string pipeName = $"RevitPipe_{processId}";

            if (!IsPipeServerRunning(pipeName))
            {
                new Thread(() =>
                {
                    while (true)
                    {
                        using (NamedPipeServerStream pipeServer = new NamedPipeServerStream(pipeName, PipeDirection.InOut))
                        {
                            pipeServer.WaitForConnection();
                            using (StreamReader reader = new StreamReader(pipeServer))
                            {
                                string elementIdStr = reader.ReadLine();
                                ElementId[] elementIds = elementIdStr
                                    .Split(',')
                                    .Select(strId => int.Parse(strId))
                                    .Select(intId => new ElementId(intId))
                                    .ToArray();

                                if (elementIds.Any())
                                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new ElementSelect(elementIds)); ;
                            }
                        }
                    }
                })
                { IsBackground = true }
                .Start();
            }
        }

        public static bool IsPipeServerRunning(string pipeName)
        {
            try
            {
                using (var client = new NamedPipeClientStream(".", pipeName, PipeDirection.In))
                {
                    client.Connect(100); // Спрабуем падключыцца (таймаут 100 мс)
                    return true;  // Калі падключыліся — сервер працуе
                }
            }
            catch (TimeoutException)
            {
                return false;
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка при проверке свободного сервера (пайпа). Отправь разработчику: {ex.Message}");
            }
        }
    }
}
