using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_ConsoleApp_Bitrix24Worker
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
