﻿namespace KPLN_Library_SQLiteWorker.Core
{
    /// <summary>
    /// Структура для описания сущности для десериализации конфига с путями к БД
    /// </summary>
    internal struct DatabasePaths
    {
        public string Name { get; set; }
        public string Path { get; set; }
    }
}
