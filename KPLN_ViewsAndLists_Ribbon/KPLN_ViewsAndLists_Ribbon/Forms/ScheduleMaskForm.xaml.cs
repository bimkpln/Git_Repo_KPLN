using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using KPLN_ViewsAndLists_Ribbon.Common.Schedules;

namespace KPLN_ViewsAndLists_Ribbon.Forms
{
    public partial class ScheduleMaskForm : Window
    {
        public sealed class MaskChoice
        {
            public string Token { get; set; }
            public string Description { get; set; }
        }

        private readonly string _sampleSource;
        private readonly bool _forName;
        private readonly bool _parameterMode;
        private readonly bool _managerMode;
        private readonly bool _allowOriginal;
        public string Pattern => PatternBox.Text;

        public ScheduleMaskForm(string pattern, bool forName = false) : this(pattern, forName, false) { }

        public ScheduleMaskForm(string pattern, bool forName, bool parameterMode) : this(pattern, forName, parameterMode, false) { }

        public ScheduleMaskForm(string pattern, bool forName, bool parameterMode, bool managerMode, bool allowOriginal = true)
        {
            _managerMode = managerMode;
            _allowOriginal = allowOriginal;
            _parameterMode = parameterMode;
            _forName = forName;
            _sampleSource = forName ? "Экспликация" : "2.-1";
            InitializeComponent();
            var choices = new List<MaskChoice>
            {
                new MaskChoice { Token = "{Этаж}", Description = "Значение столбца «Этаж» выбранного уровня. Например: 3. Маска 2.{Этаж} даст 2.3." },
                new MaskChoice { Token = "{Уровень}", Description = "Полное имя выбранного уровня Revit. Например: 03 этаж." },
                new MaskChoice { Token = "{Отметка}", Description = "Отметка уровня в метрах с тремя знаками после точки. Например: +9.000." },
                new MaskChoice { Token = "{Типовые}", Description = "Введённая вами подпись типовых этажей в таблице уровней. Например: 3–5. Если поле пустое, подставится пустой текст." },
                new MaskChoice { Token = "{Индекс}", Description = "Порядковый номер выбранного уровня, начиная с 1. Это не номер этажа; для каждой спецификации отсчёт начинается заново." },
                new MaskChoice { Token = "{Исходное}", Description = forName ? "Имя исходной спецификации, которую копируем." : "Исходное значение фильтра. Для пакетного правила у параметра должен быть ровно один исходный фильтр." }
            };
            if (forName) choices.Add(new MaskChoice
            {
                Token = "{Фильтры}",
                Description = "Индивидуальные значения итоговых фильтров через /. Фильтры с одинаковыми параметром, условием и значением во всех выбранных спецификациях исключаются. Для сравнения выберите несколько спецификаций; при одной выводятся все непустые значения. Изменения учитываются, повторы и удалённые фильтры пропускаются. Пример: Т12.1/Т22.1."
            });
            if (parameterMode)
            {
                ContextText.Text = "Для {Значение} и {Параметр} включите «По значениям параметра» на первой вкладке, выберите параметр и отметьте значения. Автоматический фильтр каждой копии использует то же значение. При обычном копировании эти две подстановки недоступны.";
                SampleText.Text = "Учебный пример: параметр «КП_О_Имя Системы», значение «Т12.1/Т22.1», индекс 1. Это не результат проверки модели.";
                PatternBox.ToolTip = "Например: {Исходное}_{Значение}. Обычный текст сохраняется, подстановки в фигурных скобках заменяются значениями.";
                choices.RemoveAll(choice => new[] { "{Этаж}", "{Уровень}", "{Отметка}", "{Типовые}" }.Contains(choice.Token));
                choices.Insert(0, new MaskChoice { Token = "{Значение}", Description = "Отмеченное значение параметра копирования с первой вкладки. Оно же применяется в автоматическом фильтре «Равно» каждой копии." });
                choices.Insert(1, new MaskChoice { Token = "{Параметр}", Description = "Название выбранного параметра копирования. Например: КП_О_Имя Системы." });
                choices.Find(choice => choice.Token == "{Индекс}").Description = "Порядковый номер создаваемой копии, начиная с 1.";
                var filters = choices.Find(choice => choice.Token == "{Фильтры}");
                if (filters != null) filters.Description = "Индивидуальные итоговые значения фильтров через /. Неизменные для всех копий значения исключаются; значение параметра копирования включается всегда.";
            }
            if (managerMode)
            {
                choices.RemoveAll(choice => choice.Token != "{Исходное}"
                    && !(forName && parameterMode && choice.Token == "{Значение}"));
                ContextText.Text = forName
                    ? "{Исходное} — имя исходной спецификации. {Значение} — выбранное значение или префикс для создаваемой копии."
                    : "{Исходное} — исходное значение этого фильтра отдельно для каждой выбранной спецификации. Например, префикс {Исходное} добавит текст перед её прежним значением. {Значение} используется только автоматической строкой при создании копий.";
                SampleText.Text = forName ? "Пример: исходная спецификация «Экспликация», значение копирования «2.1»."
                    : "Пример: исходные значения 2.1 и 2.2. Маска корпус {Исходное} даст корпус 2.1 и корпус 2.2.";
                PatternBox.ToolTip = ContextText.Text;
            }
            if (managerMode && !allowOriginal)
            {
                choices.RemoveAll(choice => choice.Token == "{Исходное}");
                ContextText.Text = "Для нового фильтра задайте конкретное значение. Маска {Исходное} доступна только для существующих фильтров.";
                SampleText.Text = "Пример: корпус 2.1 — значение нового фильтра.";
                PatternBox.ToolTip = ContextText.Text;
            }
            TokensBox.ItemsSource = choices;
            TokensBox.SelectedIndex = 0;
            PatternBox.Text = pattern ?? "";
            PatternBox.CaretIndex = PatternBox.Text.Length;
            UpdateExample(null, null);
        }

        private void SelectToken(object sender, SelectionChangedEventArgs args)
        {
            var choice = TokensBox.SelectedItem as MaskChoice;
            if (DescriptionText != null) DescriptionText.Text = choice?.Description ?? "";
        }

        private void InsertToken(object sender, RoutedEventArgs args)
        {
            var choice = TokensBox.SelectedItem as MaskChoice;
            if (choice == null) return;

            int position = PatternBox.SelectionStart;
            PatternBox.SelectedText = choice.Token;
            PatternBox.Focus();
            PatternBox.Select(position + choice.Token.Length, 0);
        }

        private void UpdateExample(object sender, TextChangedEventArgs args)
        {
            if (ExampleText == null) return;

            try
            {
                if (_managerMode) ValidateManagerPattern(Pattern, _forName && _parameterMode, _allowOriginal);
                ExampleText.Text = "Пример результата: " + ScheduleNaming.Expand(Pattern, "3", "03 этаж", "+9.000", "3–5", _sampleSource, 1, _forName ? "Т12.1/Т22.1" : null, _parameterMode ? "КП_О_Имя Системы" : null, _parameterMode ? (_managerMode ? "2.1" : "Т12.1/Т22.1") : null);
            }
            catch (ArgumentException ex) { ExampleText.Text = ex.Message; }
        }

        private void Apply(object sender, RoutedEventArgs args)
        {
            try
            {
                if (_managerMode) ValidateManagerPattern(Pattern, _forName && _parameterMode, _allowOriginal);
                if (_parameterMode && new[] { "{Этаж}", "{Уровень}", "{Отметка}", "{Типовые}" }.Any(token => Pattern.Contains(token)))
                    throw new ArgumentException("В менеджере вместо масок этажей используйте {Значение} выбранного параметра.");
                ScheduleNaming.Expand(Pattern, "3", "03 этаж", "+9.000", "3–5", _sampleSource, 1, _forName ? "Т12.1/Т22.1" : null, _parameterMode ? "КП_О_Имя Системы" : null, _parameterMode ? "Т12.1/Т22.1" : null);
                DialogResult = true;
            }
            catch (ArgumentException ex) { MessageBox.Show(this, ex.Message, Title); }
        }

        public static void ValidateManagerPattern(string pattern, bool allowValue, bool allowOriginal = true)
        {
            if (!allowOriginal && (pattern ?? "").Contains("{Исходное}"))
                throw new ArgumentException("Маска {Исходное} доступна только для существующих фильтров. Для нового фильтра задайте конкретное значение.");

            foreach (System.Text.RegularExpressions.Match match in System.Text.RegularExpressions.Regex.Matches(pattern ?? "", @"\{[^{}]*\}"))
                if (match.Value != "{Исходное}" && !(allowValue && match.Value == "{Значение}"))
                    throw new ArgumentException("Недоступная маска: " + match.Value + ". "
                        + (allowValue ? "Используйте {Исходное} или {Значение}." : allowOriginal ? "Используйте {Исходное} или конкретное значение." : "Задайте конкретное значение."));
        }
    }
}
