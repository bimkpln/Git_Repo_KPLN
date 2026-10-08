using System;
using System.Diagnostics;


namespace KPLN_Publication
{
    public class Logger
    {
        private readonly Action<string> _write;
#if Debug2026 || Revit2026
        private readonly TextWriterTraceListener _listener;
#endif

        // Для вызова без окна: без глобальных Debug.Listeners и записи рядом с DLL.
        public Logger(Action<string> write)
        {
            _write = write ?? (_ => { });
        }

        public Logger()
        {
#if !Debug2026 && !Revit2026
            Debug.Listeners.Clear();
#endif
            string assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string folder = System.IO.Path.GetDirectoryName(assemblyLocation);
            folder = System.IO.Path.Combine(folder, "logs");
            if (!System.IO.Directory.Exists(folder))
                System.IO.Directory.CreateDirectory(folder);

            string logFileName = DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".log";
            string logFilePath = System.IO.Path.Combine(folder, logFileName);
            TextWriterTraceListener tr =
                new TextWriterTraceListener(System.IO.File.CreateText(logFilePath));
#if Debug2026 || Revit2026
            // В .NET 8 нет Debug.Listeners; журнал принадлежит текущей операции.
            _listener = tr;
#else
            Debug.Listeners.Add(tr);
#endif
        }

        public void Write(string message)
        {
            if (_write != null)
            {
                _write(message);
                return;
            }

#if Debug2026 || Revit2026
            _listener.WriteLine(message);
            _listener.Flush();
#else
            Debug.WriteLine(message);
            Debug.Flush();
#endif
        }
    }
}
