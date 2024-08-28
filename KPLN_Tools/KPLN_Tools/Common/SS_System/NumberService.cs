using System;
using System.Text.RegularExpressions;

namespace KPLN_Tools.Common.SS_System
{
    internal static class NumberService
    {
        /// <summary>
        /// Метод разделения стартового номера на части
        /// </summary>
        /// <returns>Массив, где 1й элемент - стартовый номер (подвергается +1), 2й элемент - индекс системы (не меняется), 3й - разделитель</returns>
        /// <exception cref="Exception">Ошибка, если номер не парситься</exception>
        public static string[] SystemNumberSplit(string stringNumber)
        {
            // Паттерн для поиска последней цифры
            string pattern = @"([\W_]+)(\d+)$";
            Match match = Regex.Match(stringNumber, pattern);
            if (match.Success)
            {
                // Получаем символ, который отделяет части
                string separator = match.Groups[1].Value;
                // Получаем последнюю цифру
                string lastDigit = match.Groups[2].Value;
                // Все остальное до последней цифры
                string beforeLastDigit = stringNumber.Substring(0, match.Index);

                if (!int.TryParse(lastDigit, out int elemIndex))
                    throw new Exception($"Скинь разработчику: Не удалось преобразовать стартовый номер в тип int (сейчас значение - {elemIndex}).");

                return new string[] { lastDigit, beforeLastDigit, separator };
            }

            throw new Exception("Непонятный формат ввода стартового номера. Он должен всегда в конце содержать цифру, чтобы её можно было увеличивать на +1");
        }

        /// <summary>
        /// Метод сравнения значений из парамтера "КП_И_Адрес текущий"
        /// </summary>
        public static int AlphanumericCompareCurrentAdress(SS_SystemEntity x, SS_SystemEntity y)
        {
            var xAdrData = x.CurrentAdressData;
            var xSepar = x.CurrentAdressSeparator;

            var yAdrData = y.CurrentAdressData;
            var ySepar = x.CurrentAdressSeparator;

            // Разделяем строки на части
            var xParts = xAdrData.Split(xSepar.ToCharArray()); 
            var yParts = yAdrData.Split(ySepar.ToCharArray());

            int minLength = Math.Min(xParts.Length, yParts.Length);
            for (int i = 0; i < minLength; i++)
            {

                bool xIsNumber = int.TryParse(xParts[i], out int xPart);
                bool yIsNumber = int.TryParse(yParts[i], out int yPart);

                if (xIsNumber && yIsNumber)
                {
                    int comparison = xPart.CompareTo(yPart);
                    if (comparison != 0)
                        return comparison; // Если числа не равны, возвращаем результат сравнения
                }
                else
                {
                    int comparison = string.Compare(xParts[i], yParts[i], StringComparison.Ordinal);
                    if (comparison != 0)
                        return comparison; // Если строки не равны, возвращаем результат сравнения
                }
            }

            return xParts.Length.CompareTo(yParts.Length); // Сравниваем длины частей
        }
    }
}
