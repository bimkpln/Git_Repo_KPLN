using System;
using System.Linq;
using Autodesk.Revit.DB;
using F = System.Windows.Forms;

namespace KPLN_Tools_EOM.Common.ApartmentWiring
{
    internal sealed class ScopeSettings
    {
        internal bool ByWorkset;
        internal int WorksetId;
        internal string ParameterName;


        internal bool Matches(Document d, Element e)
        {
            if (ByWorkset)
                return e.WorksetId.IntegerValue == WorksetId;
            // Параметр экземпляра; при пустом значении проверяется параметр типа.
            return !string.IsNullOrWhiteSpace(Util.Source(d, e, ParameterName));
        }
    }


    internal sealed class ScopeForm : F.Form
    {
        internal ScopeSettings Settings { get; private set; }


        internal ScopeForm(Document d, string mode)
        {
            string section = mode == "Unified" ? "" : mode == "C1" ? "К1С1" : mode == "C23" ? "К1С2,3" : "К1С4";
            Text = mode == "Unified" ? "Тип отделки — область обработки" : "Отделка " + section + " — область обработки";
            Width = 760;
            Height = 330;
            StartPosition = F.FormStartPosition.CenterParent;
            var layout = new F.TableLayoutPanel
            {
                Dock = F.DockStyle.Fill,
                Padding = new F.Padding(16),
                ColumnCount = 1,
                RowCount = 7
            };
            var byParameter = new F.RadioButton
            {
                Text = "Элементы с заполненным параметром",
                Checked = true,
                AutoSize = true
            };
            var parameter = new F.ComboBox
            {
                Dock = F.DockStyle.Fill,
                DropDownStyle = F.ComboBoxStyle.DropDown
            };
            parameter.Items.AddRange(new FilteredElementCollector(d).OfClass(typeof(ParameterElement)).Cast<ParameterElement>().Select(p => p.Name).Concat(new[] { "ASML_Тип отделки", "ASML_ОВ_Тип отделки" }).Distinct().OrderBy(x => x).Cast<object>().ToArray());
            parameter.Text = "ASML_Тип отделки";
            var byWorkset = new F.RadioButton
            {
                Text = "Элементы рабочего набора",
                AutoSize = true,
                Enabled = d.IsWorkshared
            };
            var workset = new F.ComboBox
            {
                Dock = F.DockStyle.Fill,
                DropDownStyle = F.ComboBoxStyle.DropDownList,
                Enabled = false
            };
            var sets = d.IsWorkshared ? new FilteredWorksetCollector(d).OfKind(WorksetKind.UserWorkset).ToWorksets().OrderBy(w => w.Name).ToList() : null;

            if (sets != null)
            {
                workset.Items.AddRange(sets.Select(w => (object)w.Name).ToArray());

                if (sets.Count > 0)
                    workset.SelectedIndex = 0;
            }

            byParameter.CheckedChanged += (s, e) =>
            {
                parameter.Enabled = byParameter.Checked;
                workset.Enabled = !byParameter.Checked;
            };
            layout.Controls.Add(byParameter, 0, 0);
            layout.Controls.Add(parameter, 0, 1);
            layout.Controls.Add(byWorkset, 0, 2);
            layout.Controls.Add(workset, 0, 3);
            string hint = "Имя параметра можно ввести вручную. Проверяются экземпляр и тип.\n";
            hint += mode == "Unified" ? "Это только отбор: параметры записи заданы в общих настройках.\n" : mode == "C4" ? "Это только отбор: параметры записи выбраны в настройках К1С4.\n" : "Это только отбор: результат записывается в ASML_Коэффициент отделки и ASML_Тип отделки.\n";
            hint += "Изменение параметра типа влияет на все экземпляры этого типа в проекте.";
            layout.Controls.Add(new F.Label { Text = hint, AutoSize = true }, 0, 4);
            var buttons = new F.FlowLayoutPanel
            {
                Dock = F.DockStyle.Fill
            };
            var run = new F.Button
            {
                Text = "Применить",
                AutoSize = true
            };
            var cancel = new F.Button
            {
                Text = "Отмена",
                DialogResult = F.DialogResult.Cancel,
                AutoSize = true
            };
            run.Click += (s, e) =>
            {
                if (byParameter.Checked && string.IsNullOrWhiteSpace(parameter.Text))
                {
                    F.MessageBox.Show(this, "Укажите имя параметра отбора.");
                    return;
                }

                if (byWorkset.Checked && (sets == null || workset.SelectedIndex < 0))
                {
                    F.MessageBox.Show(this, "Выберите рабочий набор.");
                    return;
                }

                Settings = new ScopeSettings
                {
                    ByWorkset = byWorkset.Checked,
                    ParameterName = parameter.Text.Trim(),
                    WorksetId = byWorkset.Checked ? sets[workset.SelectedIndex].Id.IntegerValue : -1
                };
                DialogResult = F.DialogResult.OK;
                Close();
            };
            buttons.Controls.AddRange(new F.Control[] { run, cancel });
            layout.Controls.Add(buttons, 0, 5);
            Controls.Add(layout);
            AcceptButton = run;
            CancelButton = cancel;
        }
    }
}
