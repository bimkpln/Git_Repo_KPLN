using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using F = System.Windows.Forms;

namespace KPLN_Tools_EOM.Common.ApartmentWiring
{
    internal sealed class NonModelSettings
    {
        internal bool Cables = true, Fasteners = true, Copy = true, Move = true, DeleteExcess = true, ClearType = true, Overwrite = false, Preview = false, DuplicateTypes = true;
        internal string File = "", TargetFamily = "ASML_ЭОМ_Немоделируемые элементы", CabinetFamily = "ASML_ЭОМ_EKF_Щит_Встраиваемый";
        internal string Cabinet = "ASML_Принадлежность к шкафу", Quantity = "ASML_Количество", Estimate = "СМ_Смета", Section = "СМ_Секция", Floor = "СМ_Этаж";
        internal string BaseTemplate = "Тип_{type}", SumTemplate = "SUM_{type}";
        internal double Step = 2, Offset = 1, Spacing = 300;
        internal Level Level;
    }


    internal sealed class NonModelForm : F.Form
    {
        internal NonModelSettings Settings { get; private set; }

        private readonly Dictionary<string, F.Control> fields = new Dictionary<string, F.Control>();


        internal NonModelForm(Document d)
        {
            Text = "Немоделируемые элементы — этапы и настройки";
            Width = 1000;
            Height = Math.Min(810, F.Screen.PrimaryScreen.WorkingArea.Height - 60);
            StartPosition = F.FormStartPosition.CenterParent;
            var outer = new F.TableLayoutPanel
            {
                Dock = F.DockStyle.Fill,
                Padding = new F.Padding(12),
                ColumnCount = 1,
                RowCount = 3
            };
            outer.RowStyles.Add(new F.RowStyle(F.SizeType.Percent, 100));
            outer.RowStyles.Add(new F.RowStyle(F.SizeType.Absolute, 45));
            outer.RowStyles.Add(new F.RowStyle(F.SizeType.Absolute, 40));
            var tabs = new F.TabControl
            {
                Dock = F.DockStyle.Fill
            };
            var grid = TabGrid(tabs, "Общие настройки");
            var cableGrid = TabGrid(tabs, "Кабели и трубы");
            var fastenerGrid = TabGrid(tabs, "Крепления ПВХ");
            var copyGrid = TabGrid(tabs, "Секция и этаж");
            var cables = Check("Выполнять расстановку кабелей и труб", true);
            var fasteners = Check("Выполнять пересоздание креплений ПВХ", true);
            var copy = Check("Выполнять копирование секции и этажа", true);
            AddWide(cableGrid, cables, 36);
            AddWide(fastenerGrid, fasteners, 36);
            AddWide(copyGrid, copy, 36);
            Add(cableGrid, "File", "Файл Excel (.xlsx)", "");
            var browse = new F.Button
            {
                Text = "Выбрать Excel…",
                AutoSize = true
            };
            browse.Click += (sender, e) =>
            {
                using (var dialog = new F.OpenFileDialog
                {
                    Filter = "Excel (*.xlsx)|*.xlsx",
                    CheckFileExists = true
                })
                    if (dialog.ShowDialog(this) == F.DialogResult.OK)
                        fields["File"].Text = dialog.FileName;
            };
            int browseRow = cableGrid.RowCount;
            cableGrid.RowCount++;
            cableGrid.Controls.Add(browse, 1, browseRow);
            cableGrid.RowStyles.Add(new F.RowStyle(F.SizeType.Absolute, 32));
            AddFamily(grid, "TargetFamily", "Семейство немоделируемых элементов", "ASML_ЭОМ_Немоделируемые элементы", FamilyNames(d, BuiltInCategory.OST_TelephoneDevices));
            AddFamily(grid, "CabinetFamily", "Семейство щитов", "ASML_ЭОМ_EKF_Щит_Встраиваемый", FamilyNames(d, BuiltInCategory.OST_ElectricalEquipment));
            Add(grid, "Cabinet", "Параметр принадлежности к шкафу", "ASML_Принадлежность к шкафу");
            Add(grid, "Quantity", "Параметр количества", "ASML_Количество");
            Add(grid, "Estimate", "Параметр сметы", "СМ_Смета");
            Add(copyGrid, "Section", "Параметр секции", "СМ_Секция");
            Add(copyGrid, "Floor", "Параметр этажа", "СМ_Этаж");
            Add(grid, "BaseTemplate", "Шаблон обычного глобального параметра", "Тип_{type}");
            Add(grid, "SumTemplate", "Шаблон суммарного глобального параметра", "SUM_{type}");
            var levels = new FilteredElementCollector(d).OfClass(typeof(Level)).Cast<Level>().OrderBy(x => x.Elevation).ToList();
            var level = new F.ComboBox
            {
                DropDownStyle = F.ComboBoxStyle.DropDownList,
                Dock = F.DockStyle.Fill
            };
            level.Items.AddRange(levels.Select(x => (object)x.Name).ToArray());

            if (levels.Count > 0)
                level.SelectedIndex = levels.Count - 1;

            int next = cableGrid.RowCount;
            cableGrid.RowCount++;
            cableGrid.RowStyles.Add(new F.RowStyle(F.SizeType.Absolute, 32));
            cableGrid.Controls.Add(new F.Label { Text = "Уровень свободной зоны", AutoSize = true }, 0, next);
            cableGrid.Controls.Add(level, 1, next);
            var step = new F.NumericUpDown
            {
                Minimum = 0.01m,
                Maximum = 10000,
                DecimalPlaces = 2,
                Value = 2,
                Dock = F.DockStyle.Fill
            };
            var spacing = new F.NumericUpDown
            {
                Minimum = 1,
                Maximum = 100000,
                Value = 300,
                Dock = F.DockStyle.Fill
            };
            var offset = new F.NumericUpDown
            {
                Minimum = 0,
                Maximum = 10000,
                DecimalPlaces = 2,
                Value = 1,
                Dock = F.DockStyle.Fill
            };
            AddNumber(fastenerGrid, "Шаг креплений, м", step);
            AddNumber(cableGrid, "Интервал расстановки, мм", spacing);
            AddNumber(cableGrid, "Отступ свободной зоны, м", offset);
            outer.Controls.Add(tabs, 0, 0);
            var options = new F.FlowLayoutPanel
            {
                Dock = F.DockStyle.Fill
            };
            var move = Check("Перемещать существующие", true);
            var excess = Check("Удалять лишние кабели/трубы", true);
            var clear = Check("Очищать ASML_Тип", true);
            var overwrite = Check("Перезаписывать секцию и этаж", false);
            var duplicate = Check("Создавать отсутствующие типы дублированием", true);
            var preview = Check("Только проверка, без изменения модели", false);
            options.Controls.Add(preview);
            outer.Controls.Add(options, 0, 1);

            foreach (var option in new F.Control[]
            {
                move,
                excess,
                clear,
                duplicate
            })
                AddWide(cableGrid, option, 32);

            AddWide(copyGrid, overwrite, 32);
            AddWide(grid, new F.Label { AutoSize = true, Text = "НЭ: «Телефонные устройства». Щиты: «Электрооборудование».\nГруппы — из формул SUM; для секций: {section}_Тип_{type} и {section}_SUM_{type}." }, 80);
            AddWide(fastenerGrid, new F.Label { AutoSize = true, Text = "Крепления выбранного семейства удаляются и создаются заново.\nКоличество = округление вверх(количество трубы / шаг)." }, 60);
            AddWide(copyGrid, new F.Label { AutoSize = true, Text = "Секция и этаж копируются из щитов по группе шкафа." }, 45);
            AddWide(cableGrid, new F.Label { AutoSize = true, Text = "Один Excel-файл для текущей модели." }, 40);
            var buttons = new F.FlowLayoutPanel
            {
                Dock = F.DockStyle.Fill
            };
            var run = new F.Button
            {
                Text = "Запустить выбранные этапы",
                AutoSize = true
            };
            var cancel = new F.Button
            {
                Text = "Отмена",
                DialogResult = F.DialogResult.Cancel
            };
            buttons.Controls.AddRange(new F.Control[] { run, cancel });
            outer.Controls.Add(buttons, 0, 2);
            run.Click += (sender, e) =>
            {
                if (!cables.Checked && !fasteners.Checked && !copy.Checked)
                {
                    F.MessageBox.Show(this, "Выберите этап.");
                    return;
                }

                if (cables.Checked && !System.IO.File.Exists(fields["File"].Text))
                {
                    F.MessageBox.Show(this, "Выберите Excel-файл.");
                    return;
                }

                if ((cables.Checked || fasteners.Checked) && level.SelectedIndex < 0)
                {
                    F.MessageBox.Show(this, "В проекте нет уровня.");
                    return;
                }

                if (fields.Where(p => p.Key != "File").Any(p => string.IsNullOrWhiteSpace(p.Value.Text)))
                {
                    F.MessageBox.Show(this, "Выберите семейства и заполните имена и шаблоны.");
                    return;
                }

                try
                {
                    var r = new FinishRules
                    {
                        BaseTemplate = Get("BaseTemplate"),
                        SumTemplate = Get("SumTemplate"),
                        UseSections = Get("BaseTemplate").Contains("{section}")
                    };
                    r.Validate();
                }
                catch (Exception ex)
                {
                    F.MessageBox.Show(this, ex.Message);
                    return;
                }

                Settings = new NonModelSettings
                {
                    Cables = cables.Checked,
                    Fasteners = fasteners.Checked,
                    Copy = copy.Checked,
                    File = Get("File"),
                    TargetFamily = Get("TargetFamily"),
                    CabinetFamily = Get("CabinetFamily"),
                    Cabinet = Get("Cabinet"),
                    Quantity = Get("Quantity"),
                    Estimate = Get("Estimate"),
                    Section = Get("Section"),
                    Floor = Get("Floor"),
                    BaseTemplate = Get("BaseTemplate"),
                    SumTemplate = Get("SumTemplate"),
                    Step = (double)step.Value,
                    Offset = (double)offset.Value,
                    Spacing = (double)spacing.Value,
                    Level = level.SelectedIndex < 0 ? null : levels[level.SelectedIndex],
                    Move = move.Checked,
                    DeleteExcess = excess.Checked,
                    ClearType = clear.Checked,
                    Overwrite = overwrite.Checked,
                    Preview = preview.Checked,
                    DuplicateTypes = duplicate.Checked
                };
                DialogResult = F.DialogResult.OK;
                Close();
            };
            Controls.Add(outer);
            AcceptButton = run;
            CancelButton = cancel;
        }


        private static string[] FamilyNames(Document d, BuiltInCategory category)
        {
            return new FilteredElementCollector(d).OfClass(typeof(Family)).Cast<Family>().Where(x => x.FamilyCategory != null && x.FamilyCategory.Id.IntegerValue == (int)category).Select(x => x.Name).Distinct().OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase).ToArray();
        }


        private static F.TableLayoutPanel TabGrid(F.TabControl tabs, string title)
        {
            var page = new F.TabPage(title)
            {
                Padding = new F.Padding(10)
            };
            var grid = new F.TableLayoutPanel
            {
                Dock = F.DockStyle.Fill,
                ColumnCount = 2,
                AutoScroll = true
            };
            grid.ColumnStyles.Add(new F.ColumnStyle(F.SizeType.Absolute, 310));
            grid.ColumnStyles.Add(new F.ColumnStyle(F.SizeType.Percent, 100));
            page.Controls.Add(grid);
            tabs.TabPages.Add(page);
            return grid;
        }


        private static void AddWide(F.TableLayoutPanel grid, F.Control control, int height)
        {
            int row = grid.RowCount;
            grid.RowCount++;
            grid.RowStyles.Add(new F.RowStyle(F.SizeType.Absolute, height));
            grid.Controls.Add(control, 0, row);
            grid.SetColumnSpan(control, 2);
        }


        private string Get(string key)
        {
            return fields[key].Text.Trim();
        }


        private static F.CheckBox Check(string text, bool value)
        {
            return new F.CheckBox
            {
                Text = text,
                Checked = value,
                AutoSize = true
            };
        }


        private void Add(F.TableLayoutPanel g, string key, string label, string value)
        {
            int row = g.RowCount;
            g.RowCount = row + 1;
            g.RowStyles.Add(new F.RowStyle(F.SizeType.Absolute, 32));
            var input = new F.TextBox
            {
                Text = value,
                Dock = F.DockStyle.Fill
            };
            fields.Add(key, input);
            g.Controls.Add(new F.Label { Text = label, AutoSize = true }, 0, row);
            g.Controls.Add(input, 1, row);
        }


        private void AddFamily(F.TableLayoutPanel g, string key, string label, string value, string[] names)
        {
            int row = g.RowCount;
            g.RowCount++;
            g.RowStyles.Add(new F.RowStyle(F.SizeType.Absolute, 32));
            var input = new F.ComboBox
            {
                DropDownStyle = F.ComboBoxStyle.DropDownList,
                Dock = F.DockStyle.Fill,
                IntegralHeight = false,
                MaxDropDownItems = 16
            };
            input.Items.AddRange(names.Cast<object>().ToArray());
            input.SelectedIndex = Array.IndexOf(names, value);
            fields.Add(key, input);
            g.Controls.Add(new F.Label { Text = label, AutoSize = true }, 0, row);
            g.Controls.Add(input, 1, row);
        }


        private void AddNumber(F.TableLayoutPanel g, string label, F.Control input)
        {
            int row = g.RowCount;
            g.RowCount++;
            g.RowStyles.Add(new F.RowStyle(F.SizeType.Absolute, 32));
            g.Controls.Add(new F.Label { Text = label, AutoSize = true }, 0, row);
            g.Controls.Add(input, 1, row);
        }
    }
}
