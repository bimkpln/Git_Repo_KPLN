using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace KPLN_Library_Schedules.Common.Schedules
{
    public sealed class ScheduleMaskChoice
    {
        public string Token { get; set; }
        public string Description { get; set; }
    }

    /// <summary>Набор подстановок, пояснения и проверка для конкретного инструмента.</summary>
    public sealed class ScheduleMaskCatalog
    {
        public IList<ScheduleMaskChoice> Choices { get; private set; }
        public string Context { get; private set; }
        public string Sample { get; private set; }
        private bool _forName;
        private bool _allowValue;

        public static ScheduleMaskCatalog CopyByLevels(bool forName = false, bool allowOriginal = true)
        {
            var choices = new List<ScheduleMaskChoice>
            {
                Choice("{Этаж}", "Значение столбца «Этаж» выбранного уровня. Например: 3. Маска 2.{Этаж} даст 2.3."),
                Choice("{Уровень}", "Полное имя выбранного уровня Revit. Например: 03 этаж."),
                Choice("{Отметка}", "Отметка уровня в метрах с тремя знаками после точки. Например: +9.000."),
                Choice("{Типовые}", "Подпись типовых этажей из таблицы уровней. Например: 3–5. Пустая ячейка даёт пустой текст."),
                Choice("{Индекс}", "Порядковый номер выбранного уровня, начиная с 1. Это не номер этажа.")
            };
            if (allowOriginal)
                choices.Add(Choice("{Исходное}", forName ? "Имя исходной спецификации." : "Исходное значение именно этого фильтра."));

            return new ScheduleMaskCatalog
            {
                Choices = choices.AsReadOnly(), _forName = forName,
                Context = "Значения этажа, уровня, отметки и типовых этажей берутся из таблицы «Листы по этажам»."
                    + (allowOriginal ? "" : " У нового фильтра нет исходного значения."),
                Sample = "Учебный пример: этаж 3, уровень «03 этаж», отметка +9.000, типовые 3–5, индекс 1."
            };
        }

        public static ScheduleMaskCatalog Constructor(bool forName = false, bool parameterMode = false, bool allowOriginal = true)
        {
            var choices = new List<ScheduleMaskChoice>();
            if (allowOriginal)
                choices.Add(Choice("{Исходное}", forName ? "Имя исходной спецификации."
                    : "Исходное значение этого фильтра отдельно для каждой выбранной спецификации."));
            if (forName && parameterMode)
                choices.Add(Choice("{Значение}", "Выбранное значение или префикс параметра копирования с первой вкладки."));

            return new ScheduleMaskCatalog
            {
                Choices = choices.AsReadOnly(), _forName = forName, _allowValue = forName && parameterMode,
                Context = forName
                    ? "Для имени доступно {Исходное}; при копировании по параметру также {Значение}."
                    : allowOriginal
                        ? "{Исходное} — прежнее значение фильтра каждой спецификации. {Значение} задаётся только автоматическим фильтром копирования."
                        : "Для нового фильтра задайте конкретное значение. У него нет исходного значения.",
                Sample = forName ? "Пример: исходная спецификация «Экспликация», значение копирования «2.1»."
                    : allowOriginal ? "Пример: маска корпус {Исходное} для исходного значения 2.1 даст корпус 2.1."
                        : "Пример нового значения: корпус 2.1."
            };
        }

        private static ScheduleMaskChoice Choice(string token, string description)
        {
            return new ScheduleMaskChoice { Token = token, Description = description };
        }

        public void Validate(string pattern)
        {
            foreach (Match match in Regex.Matches(pattern ?? "", @"\{[^{}]*\}"))
            {
                if (Choices.Any(choice => choice.Token == match.Value))
                    continue;

                throw new ArgumentException("Недоступная маска: " + match.Value + ". "
                    + (Choices.Count == 0 ? "Задайте конкретное значение."
                        : "Доступны: " + string.Join(", ", Choices.Select(choice => choice.Token)) + "."));
            }
        }

        public string Example(string pattern)
        {
            Validate(pattern);
            return ScheduleNaming.Expand(pattern, "3", "03 этаж", "+9.000", "3–5",
                _forName ? "Экспликация" : "2.1", 1, null, null, _allowValue ? "2.1" : null);
        }
    }
}
