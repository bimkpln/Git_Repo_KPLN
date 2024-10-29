namespace KPLN_Library_ExtensibleStorage
{
    public enum MessageStatus
    {
        Ok,
        Error
    }

    /// <summary>
    /// Класс-обертка для сообщений об ExtensibleStorage
    /// </summary>
    public class ResultMessage
    {
        /// <summary>
        /// Описание, которое содержится в ExtensibleStorage БЕЗ технических пометок
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// Статус полученного ответа
        /// </summary>
        public MessageStatus CurrentStatus { get; set; }
    }
}
