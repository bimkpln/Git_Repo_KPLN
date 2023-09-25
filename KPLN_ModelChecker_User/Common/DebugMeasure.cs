using System;
using System.Diagnostics;
using System.Reflection;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_ModelChecker_User.Common
{
    /// <summary>
    /// Класс для замеров времени выполнения методов
    /// </summary>
    internal static class DebugMeasure
    {
        /// <summary>
        /// Замер без ограничений по параметрам метода без возврата результата
        /// </summary>
        /// <param name="method">Метод, для запуска</param>
        internal static void MeasureExecutionTime(Action method)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            method();

            stopwatch.Stop();
            Print($"Метод {method.Method.Name} занял: {stopwatch.Elapsed.Minutes}m:{stopwatch.Elapsed.Seconds}s:{stopwatch.Elapsed.Milliseconds}ms",
                MessageType.Warning);
        }

        /// <summary>
        /// Замер без ограничений по параметрам метода с возвратом результата
        /// </summary>
        /// <typeparam name="TResult">Тип данных результата работы метода</typeparam>
        /// <param name="method">Метод, для запуска</param>
        /// <returns>Результат работы метода</returns>
        internal static TResult MeasureExecutionTime<TResult>(Func<TResult> method)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            TResult result = method();

            stopwatch.Stop();
            Print($"Метод {method.Method.Name} занял: {stopwatch.Elapsed.Minutes}m:{stopwatch.Elapsed.Seconds}s:{stopwatch.Elapsed.Milliseconds}ms",
                MessageType.Warning);

            return result;
        }

        /// <summary>
        /// Замер с одним параметром метода с возвратом результата
        /// </summary>
        /// <typeparam name="T">Тип данных параметра метода</typeparam>
        /// <typeparam name="TResult">Тип данных результата работы метода</typeparam>
        /// <param name="n1">Переменная метода</param>
        /// <param name="method">Метод, для запуска</param>
        // <returns>Результат работы метода</returns>
        internal static TResult MeasureExecutionTime<T, TResult>(T n1, Func<T, TResult> method)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            TResult result = method(n1);

            stopwatch.Stop();
            Print($"Метод {method.Method.Name} занял: {stopwatch.Elapsed.Minutes}m:{stopwatch.Elapsed.Seconds}s:{stopwatch.Elapsed.Milliseconds}ms",
                MessageType.Warning);

            return result;
        }

        /// <summary>
        /// Замер с 2 параметром метода с возвратом результата
        /// </summary>
        /// <typeparam name="T1">Тип данных параметра метода</typeparam>
        /// <typeparam name="TResult">Тип данных результата работы метода</typeparam>
        /// <param name="n1">Переменная метода</param>
        /// <param name="method">Метод, для запуска</param>
        // <returns>Результат работы метода</returns>
        internal static TResult MeasureExecutionTime<T1, T2, TResult>(T1 n1, T2 n2, Func<T1, T2, TResult> method)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            TResult result = method(n1, n2);

            stopwatch.Stop();
            Print($"Метод {method.Method.Name} занял: {stopwatch.Elapsed.Minutes}m:{stopwatch.Elapsed.Seconds}s:{stopwatch.Elapsed.Milliseconds}ms",
                MessageType.Warning);

            return result;
        }
    }
}
