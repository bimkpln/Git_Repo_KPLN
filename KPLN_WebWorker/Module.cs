using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_WebWorker.Common;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;

namespace KPLN_WebWorker
{
    public class Module : IExternalModule
    {
        private static Process _currentWebApp;

        public Module()
        {
            ProccesId = 5100;

            IEnumerable<Process> kplnWebAppProcesses = Process
                .GetProcesses()
                .Where(pr => pr.ProcessName.ToLower().Equals("kpln_webapplication"));
            if (kplnWebAppProcesses.Any())
                _currentWebApp = kplnWebAppProcesses.FirstOrDefault();
            else
                _currentWebApp = Process.Start(@"X:\BIM\5_Scripts\Git_Repo_KPLN\KPLN_WebApplication\bin\Release\net8.0\KPLN_WebApplication.exe", ProccesId.ToString());

            PipeServerWorker.StartListening();
        }

        /// <summary>
        /// Ссылка на id процесса. Служит для увязки работы вэба и ревита
        /// </summary>
        public static int ProccesId { get; private set; }


        public Result Close()
        {
            Thread.Sleep(5000);

            IEnumerable<Process> revitProcesses = Process
                .GetProcesses()
                .Where(pr => pr.ProcessName.ToLower().Equals("revit"));

            if (revitProcesses.Count() == 1)
            {
                _currentWebApp.Kill();
                _currentWebApp.WaitForExit();
                _currentWebApp?.Dispose();
            }

            return Result.Succeeded;
        }

        public Result Execute(UIControlledApplication application, string tabName)
        {
            return Result.Succeeded;
        }
    }
}
