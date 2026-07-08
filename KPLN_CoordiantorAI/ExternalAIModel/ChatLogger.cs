using System;
using System.IO;
using System.Text;

namespace KPLN_CoordiantorAI.ExternalModel
{
    /// <summary>
    /// Класс для логирования диалогов с ИИ
    /// </summary>
    public class ChatLogger
    {
        private readonly string _logFilePath;
        private readonly string _separator = "---------------------------------------------------------------------\n---------------------------------------------------------------------";

        // Курс рубля к доллару (можно обновлять при каждом логировании или взять из API)
        private const double USD_TO_RUB_RATE = 95.0; // Актуальный курс на момент написания

        /// <summary>
        /// Конструктор. Создаёт экземпляр логгера для текущего пользователя
        /// </summary>
        public ChatLogger(string logFolder)
        {
            logFolder = (logFolder ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(logFolder))
            {
                _logFilePath = string.Empty;
                return;
            }

            // Получаем имя пользователя Windows
            string userName = Environment.UserName;

            // Создаём папку, если её нет
            if (!Directory.Exists(logFolder))
            {
                Directory.CreateDirectory(logFolder);
            }

            // Формируем путь к файлу: имя пользователя.txt
            _logFilePath = Path.Combine(logFolder, $"{userName}.txt");
        }

        /// <summary>
        /// Логирует один диалог (вопрос + ответ)
        /// </summary>
        /// <param name="question">Вопрос пользователя</param>
        /// <param name="answer">Ответ ИИ</param>
        public void Log(string question, string answer)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(_logFilePath))
                    return;

                //string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                DateTime timestamp = DateTime.Now;

                StringBuilder logEntry = new StringBuilder();
                //logEntry.AppendLine("[" + timestamp + "]");
                logEntry.AppendLine("[" + FormatLogDate(timestamp) + "]");
                AppendRequestMetadata(logEntry, timestamp, timestamp, null, null);

                logEntry.AppendLine("");
                logEntry.AppendLine($"REQUEST: {question}");
                logEntry.AppendLine($"RESPONSE: {answer}");
                logEntry.AppendLine(_separator);

                // Добавляем запись в конец файла (создаёт файл, если его нет)
                File.AppendAllText(_logFilePath, logEntry.ToString(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                // Если не удалось записать лог, не прерываем работу плагина
                System.Diagnostics.Debug.WriteLine($"Ошибка записи лога: {ex.Message}");
            }
        }



        public void Log(
            string question,
            string answer,
            DateTime requestTime,
            DateTime responseTime,
            string revitModelName,
            string revitViewName)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(_logFilePath))
                    return;

                StringBuilder logEntry = new StringBuilder();
                logEntry.AppendLine("[" + FormatLogDate(requestTime) + "]");
                AppendRequestMetadata(logEntry, requestTime, responseTime, revitModelName, revitViewName);
                logEntry.AppendLine("");
                logEntry.AppendLine($"ВОПРОС: {question}");
                logEntry.AppendLine($"ОТВЕТ: {answer}");
                logEntry.AppendLine(_separator);

                File.AppendAllText(_logFilePath, logEntry.ToString(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка записи лога: {ex.Message}");
            }
        }


        public void LogWithTokens(
            string question,
            string answer,
            int cacheHitTokens,
            int cacheMissTokens,
            int completionTokens,
            int totalTokens,
            double usdToRubRate = USD_TO_RUB_RATE)
        {
            DateTime timestamp = DateTime.Now;
            LogWithTokens(
                question,
                answer,
                cacheHitTokens,
                cacheMissTokens,
                completionTokens,
                totalTokens,
                timestamp,
                timestamp,
                null,
                null,
                usdToRubRate);
        }






        /// <summary>
        /// Логирует диалог с информацией о токенах
        /// </summary>
        public void LogWithTokens(
            string question,
            string answer,
            int cacheHitTokens,
            int cacheMissTokens,
            int completionTokens,
            int totalTokens,
            DateTime requestTime,
            DateTime responseTime,
            string revitModelName,
            string revitViewName,
            double usdToRubRate = USD_TO_RUB_RATE)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(_logFilePath))
                    return;

                //string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                string timestamp = FormatLogDate(requestTime);

                // Расчёт стоимости
                double costUSD = (cacheHitTokens * 0.0028 + cacheMissTokens * 0.14 + completionTokens * 0.28) / 1_000_000;
                double costRUB = costUSD * usdToRubRate;

                // Экономия
                double costWithoutCache = (cacheHitTokens + cacheMissTokens + completionTokens) * 0.28 / 1_000_000;
                double savingsUSD = costWithoutCache - costUSD;
                double savingsRUB = savingsUSD * usdToRubRate;

                StringBuilder logEntry = new StringBuilder();
                logEntry.AppendLine("[" + timestamp + "]");
                AppendRequestMetadata(logEntry, requestTime, responseTime, revitModelName, revitViewName);
                logEntry.AppendLine("");
                logEntry.AppendLine($"ВОПРОС: {question}");
                logEntry.AppendLine($"ОТВЕТ: {answer}");
                logEntry.AppendLine("");
                logEntry.AppendLine("--- СТАТИСТИКА ТОКЕНОВ ---");
                logEntry.AppendLine($"Промпт (Cache Hit):  {cacheHitTokens:N0} токенов");
                logEntry.AppendLine($"Промпт (Cache Miss): {cacheMissTokens:N0} токенов");
                logEntry.AppendLine($"Генерация:           {completionTokens:N0} токенов");
                logEntry.AppendLine($"ВСЕГО:               {totalTokens:N0} токенов");
                logEntry.AppendLine("");
                logEntry.AppendLine($"Стоимость:           ${costUSD:F6} (≈ {costRUB:F2} руб.)");
                logEntry.AppendLine($"Экономия:            ${savingsUSD:F6} (≈ {savingsRUB:F2} руб.)");
                logEntry.AppendLine(_separator);

                File.AppendAllText(_logFilePath, logEntry.ToString(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка записи лога токенов: {ex.Message}");
            }
        }


        /// <summary>
        /// Логирует диалог с возможностью указать пользовательский путь к файлу
        /// </summary>
        /// <param name="question">Вопрос пользователя</param>
        /// <param name="answer">Ответ ИИ</param>
        /// <param name="customFilePath">Пользовательский путь к файлу (опционально)</param>
        public void Log(string question, string answer, string customFilePath)
        {
            try
            {
                DateTime timestamp = DateTime.Now;
                string logFilePath = customFilePath ?? _logFilePath;
                if (string.IsNullOrWhiteSpace(logFilePath))
                    return;

                // Создаём папку для файла, если её нет
                string folder = Path.GetDirectoryName(logFilePath);
                if (!string.IsNullOrEmpty(folder) && !Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                }

                StringBuilder logEntry = new StringBuilder();
                //logEntry.AppendLine(timestamp);
                logEntry.AppendLine(FormatLogDate(timestamp));
                AppendRequestMetadata(logEntry, timestamp, timestamp, null, null);
                logEntry.AppendLine($"ВОПРОС: {question}");
                logEntry.AppendLine($"ОТВЕТ: {answer}");
                logEntry.AppendLine(_separator);

                File.AppendAllText(logFilePath, logEntry.ToString(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка записи лога: {ex.Message}");
            }
        }

        private static void AppendRequestMetadata(
            StringBuilder logEntry,
            DateTime requestTime,
            DateTime responseTime,
            string revitModelName,
            string revitViewName)
        {
            logEntry.AppendLine("REQUEST_TIME: " + FormatLogDate(requestTime));
            logEntry.AppendLine("RESPONSE_TIME: " + FormatLogDate(responseTime));

            if (!string.IsNullOrWhiteSpace(revitModelName))
                logEntry.AppendLine("REVIT_MODEL: " + revitModelName.Trim());

            if (!string.IsNullOrWhiteSpace(revitViewName))
                logEntry.AppendLine("REVIT_VIEW: " + revitViewName.Trim());
        }

        private static string FormatLogDate(DateTime value)
        {
            return value.ToString("yyyy-MM-dd HH:mm:ss.fff");
        }
    }
}
