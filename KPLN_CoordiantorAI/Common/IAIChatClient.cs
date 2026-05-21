using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace KPLN_CoordiantorAI.Common
{
    public interface IAiChatClient
    {
        Task<string> SendAsync(IEnumerable<ChatMessage> messages, GigaChatSettings settings, CancellationToken cancellationToken);
    }
}