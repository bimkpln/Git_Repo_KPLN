using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KPLN_Revit_Starter
{
    public class Program
    {
        [DllImport("USER32.DLL", EntryPoint = "SetForegroundWindow")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        static void Main(string[] args)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo(@"C:\Program Files\Autodesk\Revit 2023\Revit.exe")
            {
                CreateNoWindow = false,
                UseShellExecute = false,
                WindowStyle = ProcessWindowStyle.Hidden,
            };

            try
            {
                RunRevit(startInfo, args);
            }
            catch
            {
                // Log error.
            }
        }

        static void RunRevit(ProcessStartInfo startInfo, string[] sendKeysPair)
        {
            var tcs = new TaskCompletionSource<bool>();

            var process = new Process
            {
                StartInfo = startInfo,
                EnableRaisingEvents = true
            };

            process.Exited += (sender, args) =>
            {
                tcs.SetResult(true);
                process.Dispose();
            };

            process.Start();
            Thread.Sleep(30000);

            IntPtr handle = WindowHandleSearch.GetMainWindow(process.Id);
            if (SetForegroundWindow(handle))
            {
                SendKeys.SendWait(sendKeysPair[0]);
                Console.WriteLine($"sended to: {handle}");
            }
        }
    }
}
