using System.Text;

namespace KPLN_Library_ExtensibleStorage
{
    public enum MessageStatus
    {
        Ok,
        Error
    }

    /// <summary>
    /// Класс-обертка для сообщений об Extensible Storage
    /// </summary>
    public class ResultMessage
    {
        public string Description { get; set; }

        public MessageStatus CurrentStatus { get; set; }
    }
}
