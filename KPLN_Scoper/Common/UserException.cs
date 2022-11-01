using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Scoper.Common
{
    /// <summary>
    /// Пользовательское исклчение-заглушка для визуализации отображения ошибки
    /// </summary>
    internal sealed class UserException : Exception
    {
        public UserException()
        {
        }
        public UserException(string msg) : base(msg)
        {
        }
    }
}
