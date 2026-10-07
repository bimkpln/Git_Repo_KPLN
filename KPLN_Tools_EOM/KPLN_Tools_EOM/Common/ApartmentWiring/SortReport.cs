using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.Revit.UI;
using F = System.Windows.Forms;

namespace KPLN_Tools_EOM.Common.ApartmentWiring
{
    internal sealed class SortRecord
    {
        public int Id { get; set; }
        public string Category { get; set; }
        public string ElementName { get; set; }
        public string Status { get; set; }
        public string OldValue { get; set; }
        public string NewValue { get; set; }
        public string SourceValues { get; set; }
        public string Reason { get; set; }
    }


    internal sealed class SortReport
    {
        internal readonly List<SortRecord> Rows = new List<SortRecord>();
        internal readonly List<string> Notes = new List<string>();


        internal void Show(UIApplication app, SortSettings settings, string model)
        {
            var area = F.Screen.FromHandle(app.MainWindowHandle).WorkingArea;

            using (var f = new F.Form
            {
                Text = "Сортировка спецификации — подробный отчёт",
                Width = Math.Min(1350, area.Width - 40),
                Height = Math.Min(760, area.Height - 80),
                StartPosition = F.FormStartPosition.CenterParent
            })
            {
                var layout = new F.TableLayoutPanel
                {
                    Dock = F.DockStyle.Fill,
                    Padding = new F.Padding(10),
                    ColumnCount = 1,
                    RowCount = 3
                };
                layout.RowStyles.Add(new F.RowStyle(F.SizeType.Absolute, 150));
                layout.RowStyles.Add(new F.RowStyle(F.SizeType.Absolute, 38));
                layout.RowStyles.Add(new F.RowStyle(F.SizeType.Percent, 100));
                string summary = "Проект: " + model + "\r\nЦелевой параметр: " + settings.Target + "; разделитель: «" + settings.Separator + "»\r\nПорядок: " + string.Join(" → ", settings.Sources) + "\r\nОбработано: " + Rows.Count + "; " + string.Join("; ", Rows.GroupBy(r => r.Status).Select(g => g.Key + ": " + g.Count())) + "\r\n" + string.Join("\r\n", Notes);
                layout.Controls.Add(new F.TextBox { Text = summary, Dock = F.DockStyle.Fill, ReadOnly = true, Multiline = true, ScrollBars = F.ScrollBars.Vertical }, 0, 0);
                var toolbar = new F.FlowLayoutPanel
                {
                    Dock = F.DockStyle.Fill
                };
                var filter = new F.ComboBox
                {
                    DropDownStyle = F.ComboBoxStyle.DropDownList,
                    Width = 280
                };
                filter.Items.Add("Все статусы");
                filter.Items.AddRange(Rows.Select(r => r.Status).Distinct().OrderBy(x => x).Cast<object>().ToArray());
                filter.SelectedIndex = 0;
                var save = new F.Button
                {
                    Text = "Сохранить показанные строки в CSV",
                    AutoSize = true
                };
                toolbar.Controls.AddRange(new F.Control[] { filter, save });
                layout.Controls.Add(toolbar, 0, 1);
                var grid = new F.DataGridView
                {
                    Dock = F.DockStyle.Fill,
                    ReadOnly = true,
                    AllowUserToAddRows = false,
                    AllowUserToDeleteRows = false,
                    AutoGenerateColumns = false,
                    RowHeadersVisible = false,
                    AutoSizeRowsMode = F.DataGridViewAutoSizeRowsMode.None
                };
                string[] fields =
                {
                    "Id",
                    "Category",
                    "ElementName",
                    "Status",
                    "OldValue",
                    "NewValue",
                    "SourceValues",
                    "Reason"
                };
                string[] titles =
                {
                    "ID",
                    "Категория",
                    "Элемент",
                    "Статус",
                    "Старый ключ",
                    "Новый ключ",
                    "Исходные значения и их источник",
                    "Причина / пояснение"
                };
                int[] widths =
                {
                    75,
                    150,
                    180,
                    165,
                    250,
                    250,
                    520,
                    350
                };

                for (int i = 0; i < fields.Length; i++)
                    grid.Columns.Add(new F.DataGridViewTextBoxColumn { DataPropertyName = fields[i], HeaderText = titles[i], Width = widths[i] });

                Func<List<SortRecord>> visible = () => filter.SelectedIndex == 0 ? Rows : Rows.Where(r => r.Status == Convert.ToString(filter.SelectedItem)).ToList();
                filter.SelectedIndexChanged += (s, e) => grid.DataSource = visible();
                grid.DataSource = Rows;
                save.Click += (s, e) =>
                {
                    using (var dialog = new F.SaveFileDialog
                    {
                        Filter = "CSV (*.csv)|*.csv",
                        FileName = "ASML_Сортировка_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".csv",
                        AddExtension = true
                    })
                    {
                        if (dialog.ShowDialog(f) != F.DialogResult.OK)
                            return;

                        try
                        {
                            using (var writer = new StreamWriter(dialog.FileName, false, new UTF8Encoding(true)))
                            {
                                writer.WriteLine("sep=;");

                                foreach (string line in summary.Split(new[] { "\r\n" }, StringSplitOptions.None))
                                    writer.WriteLine(Escape(line));

                                writer.WriteLine(string.Join(";", titles.Select(Escape)));

                                foreach (var row in visible())
                                    writer.WriteLine(string.Join(";", new[] { row.Id.ToString(), row.Category, row.ElementName, row.Status, row.OldValue, row.NewValue, row.SourceValues, row.Reason }.Select(Escape)));
                            }

                            F.MessageBox.Show(f, "Отчёт сохранён. Строк: " + visible().Count);
                        }
                        catch (Exception ex)
                        {
                            F.MessageBox.Show(f, "Не удалось сохранить отчёт: " + ex.Message);
                        }
                    }
                };
                layout.Controls.Add(grid, 0, 2);
                f.Controls.Add(layout);
                f.ShowDialog(new WindowOwner(app));
            }
        }


        private static string Escape(string text)
        {
            text = text ?? "";
            // Экспортируем как текст, чтобы значения модели не выполнялись как формулы Excel.
            if (text.Length > 0 && "=+-@".IndexOf(text[0]) >= 0)
                text = "'" + text;

            return "\"" + text.Replace("\"", "\"\"") + "\"";
        }
    }
}
