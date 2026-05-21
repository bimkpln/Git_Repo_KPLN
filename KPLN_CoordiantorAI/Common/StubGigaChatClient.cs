using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace KPLN_CoordiantorAI.Common
{
    public sealed class StubGigaChatClient : IAiChatClient
    {
        public async Task<string> SendAsync(IEnumerable<ChatMessage> messages, GigaChatSettings settings, CancellationToken cancellationToken)
        {
            await Task.Delay(700, cancellationToken);

            ChatMessage lastUserMessage = messages == null
                ? null
                : messages.LastOrDefault(m => m != null && m.Role == ChatMessageRole.User);

            if (lastUserMessage == null || string.IsNullOrWhiteSpace(lastUserMessage.Text))
                return BuildDiagnosticMessage(settings);

            return BuildDiagnosticMessage(settings);
        }

        private static string BuildDiagnosticMessage(GigaChatSettings settings)
        {
            if (settings == null)
                return "GigaChat не запущен: настройки не загрузились из БД.";

            List<string> missingFields = new List<string>();

            if (string.IsNullOrWhiteSpace(settings.AuthUrl))
                missingFields.Add("Auth URL");
            if (string.IsNullOrWhiteSpace(settings.ApiUrl))
                missingFields.Add("API URL");
            if (string.IsNullOrWhiteSpace(settings.ClientId))
                missingFields.Add("Client ID");
            if (string.IsNullOrWhiteSpace(settings.ClientSecret))
                missingFields.Add("Client Secret");
            if (string.IsNullOrWhiteSpace(settings.Scope))
                missingFields.Add("Scope");
            if (string.IsNullOrWhiteSpace(settings.SystemPrompt))
                missingFields.Add("Системный промт");

            StringBuilder builder = new StringBuilder();

            if (missingFields.Count > 0)
            {
                builder.AppendLine("GigaChat не готов к подключению: в настройках не заполнены поля:");
                foreach (string field in missingFields)
                    builder.AppendLine("- " + field);

                builder.AppendLine();
            }
            else
            {
                builder.AppendLine("Настройки GigaChat прочитаны из БД: Client ID, Client Secret, Scope и системный промт заполнены.");
                builder.AppendLine();
            }

            builder.AppendLine("Реальный вызов GigaChat API пока не подключен: сейчас команда старта использует диагностическую заглушку StubGigaChatClient.");
            builder.AppendLine("Поэтому токены уже можно сохранять и проверять, но запрос в GigaChat ещё не отправляется.");

            return builder.ToString().TrimEnd();
        }
    }
}