using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using System.Text;
using Autodesk.Revit.DB;
using F = System.Windows.Forms;

namespace KPLN_Tools_EOM.Common.ApartmentWiring
{
    [DataContract]
    internal sealed class SortSettings
    {
        [DataMember]
        public string Target = "ASML_Сортировка спецификации";
        [DataMember]
        public string Separator = "/";
        [DataMember]
        public List<string> Sources = new List<string>
        {
            "ASML_Раздел спецификации",
            "ASML_Наименование",
            "ASML_Тип",
            "ASML_Код изделия"
        };
        [DataMember]
        public bool Remember = true;


        private static string PathFor(Document d)
        {
            string key = string.IsNullOrEmpty(d.PathName) ? d.ProjectInformation.UniqueId : d.PathName;
            string hash;

            using (var sha = SHA256.Create())
                hash = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(key))).Replace("-", "");

            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ASML", "Revit2020", "Sort-" + hash + ".json");
        }


        internal static SortSettings Load(Document d)
        {
            if (!File.Exists(PathFor(d)))
                return new SortSettings();

            try
            {
                using (var stream = File.OpenRead(PathFor(d)))
                {
                    var s = (SortSettings)new DataContractJsonSerializer(typeof(SortSettings)).ReadObject(stream);

                    if (s == null || string.IsNullOrWhiteSpace(s.Target) || s.Separator == null || s.Sources == null || s.Sources.Count == 0 || s.Sources.Any(string.IsNullOrWhiteSpace))
                        throw new InvalidDataException();

                    return s;
                }
            }
            catch
            {
                F.MessageBox.Show("Настройки сортировки повреждены. Загружены исходные параметры.");
                return new SortSettings();
            }
        }


        internal void Save(Document d)
        {
            string p = PathFor(d);
            Directory.CreateDirectory(Path.GetDirectoryName(p));

            using (var stream = File.Create(p))
                new DataContractJsonSerializer(typeof(SortSettings)).WriteObject(stream, this);
        }
    }


    internal sealed class SortForm : F.Form
    {
        internal SortSettings Settings { get; private set; }

        private readonly F.DataGridView sources;


        internal SortForm(Document d, SortSettings s)
        {
            Text = "Сортировка спецификации — параметры ключа";
            Width = 900;
            Height = 640;
            MinimumSize = new System.Drawing.Size(760, 550);
            StartPosition = F.FormStartPosition.CenterParent;
            var layout = new F.TableLayoutPanel
            {
                Dock = F.DockStyle.Fill,
                Padding = new F.Padding(12),
                ColumnCount = 1,
                RowCount = 6
            };
            layout.RowStyles.Add(new F.RowStyle(F.SizeType.Absolute, 80));
            layout.RowStyles.Add(new F.RowStyle(F.SizeType.Absolute, 38));
            layout.RowStyles.Add(new F.RowStyle(F.SizeType.Percent, 100));
            layout.RowStyles.Add(new F.RowStyle(F.SizeType.Absolute, 38));
            layout.RowStyles.Add(new F.RowStyle(F.SizeType.Absolute, 65));
            layout.RowStyles.Add(new F.RowStyle(F.SizeType.Absolute, 38));
            var header = new F.TableLayoutPanel
            {
                Dock = F.DockStyle.Fill,
                ColumnCount = 2
            };
            header.ColumnStyles.Add(new F.ColumnStyle(F.SizeType.Absolute, 210));
            header.ColumnStyles.Add(new F.ColumnStyle(F.SizeType.Percent, 100));
            var target = new F.TextBox
            {
                Text = s.Target,
                Dock = F.DockStyle.Fill
            };
            var separator = new F.TextBox
            {
                Text = s.Separator,
                Dock = F.DockStyle.Fill
            };
            header.Controls.Add(new F.Label { Text = "Куда записать ключ", AutoSize = true }, 0, 0);
            header.Controls.Add(target, 1, 0);
            header.Controls.Add(new F.Label { Text = "Разделитель частей", AutoSize = true }, 0, 1);
            header.Controls.Add(separator, 1, 1);
            layout.Controls.Add(header, 0, 0);
            var addPanel = new F.FlowLayoutPanel
            {
                Dock = F.DockStyle.Fill
            };
            var names = new F.ComboBox
            {
                Width = 630,
                DropDownStyle = F.ComboBoxStyle.DropDown
            };
            names.Items.AddRange(new FilteredElementCollector(d).OfClass(typeof(ParameterElement)).Cast<ParameterElement>().Select(p => p.Name).Concat(s.Sources).Distinct().OrderBy(n => n).Cast<object>().ToArray());
            var add = new F.Button
            {
                Text = "Добавить",
                AutoSize = true
            };
            addPanel.Controls.AddRange(new F.Control[] { names, add });
            layout.Controls.Add(addPanel, 0, 1);
            sources = new F.DataGridView
            {
                Dock = F.DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = true,
                MultiSelect = false,
                SelectionMode = F.DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = F.DataGridViewAutoSizeColumnsMode.Fill
            };
            sources.Columns.Add("Parameter", "Исходный параметр — порядок сверху вниз");

            foreach (string name in s.Sources)
                sources.Rows.Add(name);

            add.Click += (sender, e) =>
            {
                if (!string.IsNullOrWhiteSpace(names.Text))
                    sources.Rows.Add(names.Text.Trim());
            };
            layout.Controls.Add(sources, 0, 2);
            var order = new F.FlowLayoutPanel
            {
                Dock = F.DockStyle.Fill
            };
            var up = new F.Button
            {
                Text = "Выше"
            };
            var down = new F.Button
            {
                Text = "Ниже"
            };
            var remove = new F.Button
            {
                Text = "Удалить"
            };
            up.Click += (sender, e) => MoveSelectedRow(-1);
            down.Click += (sender, e) => MoveSelectedRow(1);
            remove.Click += (sender, e) =>
            {
                if (sources.CurrentRow != null)
                    sources.Rows.RemoveAt(sources.CurrentRow.Index);
            };
            order.Controls.AddRange(new F.Control[] { up, down, remove });
            layout.Controls.Add(order, 0, 3);
            layout.Controls.Add(new F.Label { AutoSize = true, Text = "Имя можно выбрать из списка или ввести вручную; строки таблицы можно редактировать.\nЗначение берётся у экземпляра, при пустом значении — у типа. Пустые части сохраняют место в ключе.\nЕсли все части пусты, прежний ключ очищается. Запись — только в текстовый параметр экземпляра." }, 0, 4);
            var buttons = new F.FlowLayoutPanel
            {
                Dock = F.DockStyle.Fill
            };
            var remember = new F.CheckBox
            {
                Text = "Запомнить для проекта",
                Checked = s.Remember,
                AutoSize = true
            };
            var run = new F.Button
            {
                Text = "Запустить",
                AutoSize = true
            };
            var cancel = new F.Button
            {
                Text = "Отмена",
                DialogResult = F.DialogResult.Cancel
            };
            run.Click += (sender, e) =>
            {
                sources.EndEdit();
                var list = sources.Rows.Cast<F.DataGridViewRow>().Select(r => Convert.ToString(r.Cells[0].Value).Trim()).ToList();

                if (string.IsNullOrWhiteSpace(target.Text) || list.Count == 0 || list.Any(string.IsNullOrWhiteSpace))
                {
                    F.MessageBox.Show(this, "Укажите целевой параметр и хотя бы один заполненный исходный параметр.");
                    return;
                }

                if (list.Distinct(StringComparer.Ordinal).Count() != list.Count)
                {
                    F.MessageBox.Show(this, "Исходные параметры повторяются.");
                    return;
                }

                if (list.Contains(target.Text.Trim()))
                {
                    F.MessageBox.Show(this, "Целевой параметр не должен входить в исходные: это изменит ключ при повторном запуске.");
                    return;
                }

                Settings = new SortSettings
                {
                    Target = target.Text.Trim(),
                    Separator = separator.Text,
                    Sources = list,
                    Remember = remember.Checked
                };
                DialogResult = F.DialogResult.OK;
                Close();
            };
            buttons.Controls.AddRange(new F.Control[] { remember, run, cancel });
            layout.Controls.Add(buttons, 0, 5);
            Controls.Add(layout);
            AcceptButton = run;
            CancelButton = cancel;
        }


        private void MoveSelectedRow(int delta)
        {
            sources.EndEdit();

            if (sources.CurrentRow == null)
                return;

            int i = sources.CurrentRow.Index, j = i + delta;

            if (j < 0 || j >= sources.Rows.Count)
                return;

            object temp = sources.Rows[i].Cells[0].Value;
            sources.Rows[i].Cells[0].Value = sources.Rows[j].Cells[0].Value;
            sources.Rows[j].Cells[0].Value = temp;
            sources.CurrentCell = sources.Rows[j].Cells[0];
        }
    }
}
