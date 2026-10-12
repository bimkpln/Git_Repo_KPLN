using System;

namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    internal sealed class OperationError
    {
        internal string Summary, Details;

        internal static string Message(Exception error)
        {
            for (Exception current = error; current != null; current = current.InnerException)
                if (!string.IsNullOrWhiteSpace(current.Message))
                    return string.Join(" ", current.Message.Split((char[])null, StringSplitOptions.RemoveEmptyEntries));
            return error.GetType().FullName;
        }

        internal static OperationError Create(string stage, Exception error, string savedPath, bool writeAttempt = true)
        {
            string saved = savedPath == null ? "" : "Копия сохранена. ";
            return new OperationError
            {
                Summary = saved + stage + ": " + Message(error),
                Details = "Этап: " + stage + "\n" + (savedPath == null ? (writeAttempt ? "Копия не сохранена.\n" : "Сохранение файла не выполнялось.\n")
                    : "Копия сохранена: " + savedPath + "\n") + "\n" + error.ToString()
            };
        }
    }
}
