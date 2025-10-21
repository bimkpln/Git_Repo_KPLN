namespace KPLN_Loader.Core
{
    /// <summary>
    /// Структура для описания сущности для десериализации конфига с данными по вебхукам Bitrix
    /// </summary>
    public sealed class Bitrix_Config
    {
        public string Name { get; set; }

        public string URL { get; set; }
    }
}
