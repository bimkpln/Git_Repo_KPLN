using System;

namespace KPLN_Library_OpenDocHandler.Core
{
    /// <summary>
    /// Итерфейс для привязки пользовательского FieldChangedEventArgs
    /// </summary>
    public interface IFieldChangedNotifier
    {
        /// <summary>
        /// Событие, которое возникает при изменении поля для передачи
        /// </summary>
        event EventHandler<FieldChangedEventArgs> FieldChanged;
    }
}
