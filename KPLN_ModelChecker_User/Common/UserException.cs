using System;

namespace KPLN_ModelChecker_User.Common
{
    /// <summary>
    /// Пользовательская ошибка для отлова некорректного поведения пользователя
    /// </summary>
    internal class UserException : Exception
    {
        public UserException()
        {
        }

        public UserException(string message) : base(message)
        {
        }

        public UserException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
