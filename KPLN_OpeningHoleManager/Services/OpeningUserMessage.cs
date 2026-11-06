using System;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace KPLN_OpeningHoleManager.Services
{
    internal static class OpeningUserMessage
    {
        internal static string FromException(Exception exception)
        {
            WriteDiagnostic(exception?.ToString());
            while (exception is TargetInvocationException && exception.InnerException != null)
                exception = exception.InnerException;

            string message = Translate(exception?.Message);
            if (message != null)
                return message;

            if (exception?.GetType().Name == "OperationCanceledException")
                return "Операция отменена.";
            if (exception is UnauthorizedAccessException)
                return "Недостаточно прав для выполнения операции. Проверьте доступ к модели или файлу.";
            if (exception is FileNotFoundException || exception is DirectoryNotFoundException)
                return "Не найден необходимый файл или папка. Проверьте путь и доступ к сетевому диску.";
            if (exception is IOException)
                return "Не удалось прочитать или записать файл. Проверьте доступ к папке и убедитесь, что файл не занят другой программой.";
            if (exception is System.Xml.XmlException || exception?.GetType().Name.StartsWith("Json", StringComparison.Ordinal) == true)
                return "Сохранённые данные имеют неверный формат. Передайте журнал KPLN_OHM разработчику.";

            return "Не удалось выполнить операцию. Повторите попытку. Если ошибка повторится, передайте журнал KPLN_OHM разработчику.";
        }

        internal static string FromFailure(string description)
        {
            WriteDiagnostic(description);
            return Translate(description)
                ?? "Revit не разрешил сохранить изменения. Проверьте указанные элементы. Подробности записаны в журнал KPLN_OHM.";
        }

        private static string Translate(string text)
        {
            string message = (text ?? string.Empty).Trim();
            if (Contains(message, "Setting active view is temporarily disabled"))
                return "Revit временно запрещает переключение вида. Завершите текущую команду или закройте диалоговое окно, затем повторите «Показать». Можно вручную открыть 3D-вид и повторить команду.";
            if (Contains(message, "Cannot re-enter the pick operation"))
                return "Уже выполняется выбор элементов. Завершите его или нажмите Esc, затем повторите команду.";
            if (Contains(message, "owned by another user") || Contains(message, "owned by other user"))
                return "Элемент занят другим пользователем. Попросите освободить его и повторите операцию.";
            if (Contains(message, "not editable") || Contains(message, "not modifiable") || Contains(message, "read-only"))
                return "Модель или элемент недоступны для изменения. Проверьте права редактирования и завершите текущую команду.";
            if (Contains(message, "Modification of the document is forbidden") || Contains(message, "outside of transaction"))
                return "Revit сейчас запрещает изменение модели. Завершите текущую команду и повторите операцию.";

            // Собственные русские сообщения сохраняем вместе с именами моделей, типов и идентификаторами.
            // Английское сообщение с русским именем файла не считаем русской инструкцией.
            if (Regex.IsMatch(message, @"^(?:Revit\s+)?[«(]*[А-Яа-яЁё]"))
                return message;

            return null;
        }

        private static bool Contains(string message, string fragment) =>
            message.IndexOf(fragment, StringComparison.OrdinalIgnoreCase) >= 0;

        private static void WriteDiagnostic(string detail)
        {
            if (string.IsNullOrWhiteSpace(detail))
                return;

            try
            {
                string directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "KPLN", "OpeningHoleManager", "Logs");
                Directory.CreateDirectory(directory);
                string path = Path.Combine(directory, "KPLN_OHM_" + DateTime.Now.ToString("yyyy-MM-dd") + ".log");
                File.AppendAllText(path, DateTimeOffset.Now.ToString("O") + Environment.NewLine + detail + Environment.NewLine + Environment.NewLine, Encoding.UTF8);
            }
            catch
            {
                // Ошибка журнала не должна мешать показать пользователю основное сообщение.
            }
        }
    }
}
