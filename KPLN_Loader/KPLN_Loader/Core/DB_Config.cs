namespace KPLN_Loader.Core
{
    /// <summary>
    /// Структура для описания сущности для десериализации конфига с путями к БД
    /// </summary>
    public sealed class DB_Config
    {
        public string Name { get; set; }
        
        public string Path { get; set; }
        
        public string Description { get; set; }
    }
}
