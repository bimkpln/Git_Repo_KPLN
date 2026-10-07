using System;
using System.Collections.Generic;
using System.Drawing;
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
    internal sealed class GlobalForm : F.Form
    {
        private readonly F.TextBox path = new F.TextBox
        {
            ReadOnly = true,
            Dock = F.DockStyle.Fill
        };
        private readonly F.CheckBox update = new F.CheckBox
        {
            Text = "Обновлять существующие параметры",
            Checked = true,
            AutoSize = true
        };
        internal string Path
        {
            get
            {
                return path.Text;
            }
        }

        internal bool UpdateExisting
        {
            get
            {
                return update.Checked;
            }
        }


        internal GlobalForm()
        {
            Text = "Глобальные параметры из Excel";
            Width = 750;
            Height = 240;
            StartPosition = F.FormStartPosition.CenterParent;
            var grid = new F.TableLayoutPanel
            {
                Dock = F.DockStyle.Fill,
                Padding = new F.Padding(12),
                ColumnCount = 2,
                RowCount = 4
            };
            grid.ColumnStyles.Add(new F.ColumnStyle(F.SizeType.Percent, 100));
            grid.ColumnStyles.Add(new F.ColumnStyle(F.SizeType.Absolute, 110));
            grid.Controls.Add(new F.Label { Text = "Лист «Глобальные параметры»: A — имя, B — число, C — формула Revit.", AutoSize = true }, 0, 0);
            grid.Controls.Add(path, 0, 1);
            var browse = new F.Button
            {
                Text = "Выбрать…",
                Dock = F.DockStyle.Fill
            };
            browse.Click += (s, e) =>
            {
                using (var dialog = new F.OpenFileDialog
                {
                    Filter = "Excel (*.xlsx)|*.xlsx",
                    CheckFileExists = true,
                    InitialDirectory = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(typeof(GlobalCommand).Assembly.Location), "Templates")
                })
                    if (dialog.ShowDialog(this) == F.DialogResult.OK)
                        path.Text = dialog.FileName;
            };
            grid.Controls.Add(browse, 1, 1);
            grid.Controls.Add(update, 0, 2);
            var run = new F.Button
            {
                Text = "Импортировать",
                AutoSize = true
            };
            run.Click += (s, e) =>
            {
                if (!File.Exists(path.Text))
                {
                    F.MessageBox.Show(this, "Выберите существующий XLSX-файл.");
                    return;
                }

                DialogResult = F.DialogResult.OK;
                Close();
            };
            var cancel = new F.Button
            {
                Text = "Отмена",
                DialogResult = F.DialogResult.Cancel,
                AutoSize = true
            };
            grid.Controls.Add(run, 0, 3);
            grid.Controls.Add(cancel, 1, 3);
            Controls.Add(grid);
            AcceptButton = run;
            CancelButton = cancel;
        }
    }
}
