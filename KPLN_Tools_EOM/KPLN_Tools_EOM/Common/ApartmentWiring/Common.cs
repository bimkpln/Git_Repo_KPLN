using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using WinForms = System.Windows.Forms;

namespace KPLN_Tools_EOM.Common.ApartmentWiring
{
    internal sealed class WindowOwner : WinForms.IWin32Window
    {
        public IntPtr Handle { get; private set; }


        public WindowOwner(UIApplication app)
        {
            Handle = app.MainWindowHandle;
        }
    }


    internal sealed class Report
    {
        internal readonly List<string> Lines = new List<string>();
        internal int Changed, Skipped;


        internal void Show(UIApplication app, string title)
        {
            using (var f = new WinForms.Form
            {
                Text = title,
                Width = 900,
                Height = 600,
                StartPosition = WinForms.FormStartPosition.CenterParent
            })
            {
                var text = new WinForms.TextBox
                {
                    Dock = WinForms.DockStyle.Fill,
                    Multiline = true,
                    ReadOnly = true,
                    ScrollBars = WinForms.ScrollBars.Both,
                    WordWrap = false
                };
                text.Text = "Изменений: " + Changed + "; пропущено: " + Skipped + Environment.NewLine + string.Join(Environment.NewLine, Lines);
                var save = new WinForms.Button
                {
                    Text = "Сохранить отчёт в TXT",
                    Dock = WinForms.DockStyle.Bottom,
                    Height = 34
                };
                save.Click += (sender, args) =>
                {
                    using (var dialog = new WinForms.SaveFileDialog
                    {
                        Filter = "Текстовый отчёт (*.txt)|*.txt",
                        FileName = "ASML_Report_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".txt"
                    })
                    {
                        if (dialog.ShowDialog(f) != WinForms.DialogResult.OK)
                            return;

                        try
                        {
                            System.IO.File.WriteAllText(dialog.FileName, text.Text, new System.Text.UTF8Encoding(true));
                        }
                        catch (Exception ex)
                        {
                            WinForms.MessageBox.Show(f, "Отчёт не сохранён: " + ex.Message);
                        }
                    }
                };
                f.Controls.Add(text);
                f.Controls.Add(save);
                f.ShowDialog(new WindowOwner(app));
            }
        }
    }


    internal static class Util
    {
        internal static bool IsSetunProject(ExternalCommandData c)
        {
            var document = c.Application.ActiveUIDocument?.Document;
            string fileName = document == null ? "" : document.Title;

            if (fileName.StartsWith("СЕТ_1", StringComparison.Ordinal))
                return true;

            TaskDialog.Show("KPLN: Сетунь", "Работа остановлена, т.к. это плагин только для проекта «Сетунь».\nИмя файла должно начинаться с «СЕТ_1».");
            return false;
        }


        internal static Document Project(ExternalCommandData c)
        {
            var d = c.Application.ActiveUIDocument == null ? null : c.Application.ActiveUIDocument.Document;

            if (d == null || d.IsFamilyDocument || d.IsReadOnly)
                throw new InvalidOperationException("Откройте доступный для изменения проект Revit.");

            return d;
        }


        internal static Dictionary<string, GlobalParameter> Globals(Document d)
        {
            return GlobalParametersManager.GetAllGlobalParameters(d).Select(id => (GlobalParameter)d.GetElement(id)).ToDictionary(p => p.Name, StringComparer.Ordinal);
        }


        internal static Parameter Unique(Element e, string name)
        {
            if (e == null)
                return null;

            var ps = e.GetParameters(name);

            if (ps.Count > 1)
                throw new InvalidOperationException("Несколько параметров с именем «" + name + "». Требуется однозначное имя.");

            return ps.Count == 0 ? null : ps[0];
        }


        internal static Parameter InstanceOrType(Document d, Element e, string name)
        {
            return Unique(e, name) ?? Unique(d.GetElement(e.GetTypeId()), name);
        }


        internal static string Text(Document d, Parameter p)
        {
            if (p == null || !p.HasValue)
                return "";

            string value;

            switch (p.StorageType)
            {
                case StorageType.String:
                    value = p.AsString();
                    break;
                case StorageType.Integer:
                    value = p.AsValueString() ?? p.AsInteger().ToString(CultureInfo.InvariantCulture);
                    break;
                case StorageType.Double:
                    value = p.AsValueString() ?? p.AsDouble().ToString(CultureInfo.InvariantCulture);
                    break;
                case StorageType.ElementId:
                    var id = p.AsElementId();
                    value = id.IntegerValue < 0 ? "" : d.GetElement(id)?.Name ?? id.IntegerValue.ToString();
                    break;
                default:
                    value = "";
                    break;
            }

            return (value ?? "").Trim();
        }


        internal static string Source(Document d, Element e, string name)
        {
            var value = Text(d, Unique(e, name));
            return value.Length > 0 ? value : Text(d, Unique(d.GetElement(e.GetTypeId()), name));
        }


        internal static void Commit(Transaction t)
        {
            if (t.Commit() != TransactionStatus.Committed)
                throw new InvalidOperationException("Revit не подтвердил транзакцию. Изменения не сохранены.");
        }


        internal static Result Fail(Exception ex, ref string message)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }
}
