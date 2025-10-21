namespace KPLN_Library_SQLiteWorker.Core
{
    /// <summary>
    /// Структура для описания сущности для десериализации конфига с путями к БД
    /// </summary>
    internal sealed class DatabasePaths
    {
        public string Name { get; set; }
        
        public string Path { get; set; }
        
        public string Description { get; set; }
    }
}
