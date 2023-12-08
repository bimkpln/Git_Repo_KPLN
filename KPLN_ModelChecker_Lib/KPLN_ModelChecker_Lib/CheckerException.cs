using System;

namespace KPLN_ModelChecker_Lib
{
    /// <summary>
    /// Пользовательская ошибка для отлова некорректного поведения пользователя
    /// </summary>
    public class CheckerException : Exception
    {
        public CheckerException()
        {
        }

        public CheckerException(string message) : base(message)
        {
        }

        public CheckerException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
