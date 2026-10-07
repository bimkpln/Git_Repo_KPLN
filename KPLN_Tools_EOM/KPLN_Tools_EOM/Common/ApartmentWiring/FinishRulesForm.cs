using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using F = System.Windows.Forms;

namespace KPLN_Tools_EOM.Common.ApartmentWiring
{
    internal sealed class FinishRulesForm : F.Form
    {
        internal FinishRules Settings { get; private set; }

        private readonly Dictionary<string, F.ComboBox> fields = new Dictionary<string, F.ComboBox>();


        internal FinishRulesForm(Document d, FinishRules s)
        {
            Text = "Тип отделки — общие правила";
            Width = 980;
            Height = 700;
            StartPosition = F.FormStartPosition.CenterParent;
            var layout = new F.TableLayoutPanel
            {
                Dock = F.DockStyle.Fill,
                Padding = new F.Padding(12),
                ColumnCount = 2,
                RowCount = 14
            };
            layout.ColumnStyles.Add(new F.ColumnStyle(F.SizeType.Absolute, 300));
            layout.ColumnStyles.Add(new F.ColumnStyle(F.SizeType.Percent, 100));
            string[] parameterNames = new FilteredElementCollector(d).OfClass(typeof(ParameterElement)).Cast<ParameterElement>().Select(p => p.Name).Distinct().OrderBy(x => x).ToArray();
            Add(layout, "Source", "Исходный параметр шкафа", s.Source, parameterNames);
            Add(layout, "Target", "Параметр коэффициента (связь)", s.Target, parameterNames);
            Add(layout, "Finish", "Параметр текста отделки", s.Finish, parameterNames);
            Add(layout, "Marker", "Маркер шкафа", s.Marker, null);
            Add(layout, "SectionMarker", "Маркер секции", s.SectionMarker, null);
            Add(layout, "AllowedSections", "Допустимые секции (пусто = все)", s.AllowedSections, null);
            Add(layout, "BaseTemplate", "Шаблон обычного глобального параметра", s.BaseTemplate, null);
            Add(layout, "SumTemplate", "Шаблон суммарного глобального параметра", s.SumTemplate, null);
            Add(layout, "ExcludeFamily", "Исключить семейства по подстроке", s.ExcludeFamily, null);
            var checks = new F.FlowLayoutPanel
            {
                Dock = F.DockStyle.Fill,
                AutoSize = true
            };
            var sections = new F.CheckBox
            {
                Text = "Использовать секции",
                Checked = s.UseSections,
                AutoSize = true
            };
            var sum = new F.CheckBox
            {
                Text = "Искать SUM в первую очередь",
                Checked = s.UseSum,
                AutoSize = true
            };
            var fallback = new F.CheckBox
            {
                Text = "Искать обычный параметр",
                Checked = s.Fallback,
                AutoSize = true
            };
            var sectionText = new F.CheckBox
            {
                Text = "Добавлять секцию в текст отделки",
                Checked = s.SectionInText,
                AutoSize = true
            };
            var remember = new F.CheckBox
            {
                Text = "Запомнить правила для проекта",
                Checked = s.Remember,
                AutoSize = true
            };
            checks.Controls.AddRange(new F.Control[] { sections, sum, fallback, sectionText, remember });
            layout.Controls.Add(checks, 0, 9);
            layout.SetColumnSpan(checks, 2);
            layout.RowStyles.Add(new F.RowStyle(F.SizeType.Absolute, 80));
            sections.CheckedChanged += (sender, e) =>
            {
                if (sections.Checked)
                {
                    if (fields["BaseTemplate"].Text == "Тип_{type}")
                        fields["BaseTemplate"].Text = "{section}_Тип_{type}";

                    if (fields["SumTemplate"].Text == "SUM_{type}")
                        fields["SumTemplate"].Text = "{section}_SUM_{type}";
                }
                else
                {
                    if (fields["BaseTemplate"].Text == "{section}_Тип_{type}")
                        fields["BaseTemplate"].Text = "Тип_{type}";

                    if (fields["SumTemplate"].Text == "{section}_SUM_{type}")
                        fields["SumTemplate"].Text = "SUM_{type}";
                }
            };
            var help = new F.Label
            {
                AutoSize = true,
                Text = "{type} — номер типа, например 1.2. {section} — секция, например С2.\nБез секций: ЩК1.2 → Тип_{type} / SUM_{type}. С секциями: С2ЩК1.2 → {section}_Тип_{type}.\nДопустимые секции: С2, С3. Маркер С распознаёт кириллическую С и латинскую C.\nСостав SUM читается из формулы: Тип_1 + Тип_1.6 + Тип_1.7 → Отделка типов 1, 1.6, 1.7.\nПоддерживается сумма обычных типов. Условия, множители и вложенные SUM дают сообщение в отчёте."
            };
            layout.Controls.Add(help, 0, 10);
            layout.SetColumnSpan(help, 2);
            layout.RowStyles.Add(new F.RowStyle(F.SizeType.Absolute, 110));
            var buttons = new F.FlowLayoutPanel
            {
                Dock = F.DockStyle.Fill
            };
            var next = new F.Button
            {
                Text = "Далее: область обработки",
                AutoSize = true
            };
            var cancel = new F.Button
            {
                Text = "Отмена",
                DialogResult = F.DialogResult.Cancel
            };
            next.Click += (sender, e) =>
            {
                var rules = new FinishRules
                {
                    Source = Get("Source"),
                    Target = Get("Target"),
                    Finish = Get("Finish"),
                    Marker = Get("Marker"),
                    SectionMarker = Get("SectionMarker"),
                    AllowedSections = Get("AllowedSections"),
                    BaseTemplate = Get("BaseTemplate"),
                    SumTemplate = Get("SumTemplate"),
                    ExcludeFamily = Get("ExcludeFamily"),
                    UseSections = sections.Checked,
                    UseSum = sum.Checked,
                    Fallback = fallback.Checked,
                    SectionInText = sectionText.Checked,
                    Remember = remember.Checked
                };

                try
                {
                    rules.Validate();
                    Settings = rules;
                    DialogResult = F.DialogResult.OK;
                    Close();
                }
                catch (Exception ex)
                {
                    F.MessageBox.Show(this, ex.Message);
                }
            };
            buttons.Controls.AddRange(new F.Control[] { next, cancel });
            layout.Controls.Add(buttons, 0, 11);
            layout.SetColumnSpan(buttons, 2);
            layout.RowStyles.Add(new F.RowStyle(F.SizeType.Absolute, 40));
            Controls.Add(layout);
            AcceptButton = next;
            CancelButton = cancel;
        }


        private string Get(string name)
        {
            return fields[name].Text.Trim();
        }


        private void Add(F.TableLayoutPanel layout, string key, string label, string value, string[] options)
        {
            int row = fields.Count;
            layout.RowStyles.Add(new F.RowStyle(F.SizeType.Absolute, 34));
            var input = new F.ComboBox
            {
                Text = value,
                Dock = F.DockStyle.Fill,
                DropDownStyle = F.ComboBoxStyle.DropDown
            };

            if (options != null)
                input.Items.AddRange(options.Cast<object>().ToArray());

            fields.Add(key, input);
            layout.Controls.Add(new F.Label { Text = label, AutoSize = true }, 0, row);
            layout.Controls.Add(input, 1, row);
        }
    }
}
