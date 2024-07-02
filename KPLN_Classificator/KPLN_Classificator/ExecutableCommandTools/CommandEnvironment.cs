namespace KPLN_Classificator.ExecutableCommand
{
    /// <summary>
    /// Интерфейс, определяющий среду для запуска команд Revit
    /// </summary>
    public interface CommandEnvironment
    {
        /// <summary>
        /// Метод передаёт объект реализующий интерфейс среды запуска в соответствующую очередь Queue для запуска команды в событии OnIdling
        /// </summary>
        /// <param name="obj"></param>
        void toEnqueue(object obj);
    }
}
