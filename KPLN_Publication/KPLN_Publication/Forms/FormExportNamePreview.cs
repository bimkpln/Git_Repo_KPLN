using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace KPLN_Publication
{
    internal sealed class ExportNameRow
    {
        public string Format { get; set; }
        public string Model { get; set; }
        public string Sheet { get; set; }
        public string FileName { get; set; }
        public string Status { get; set; }
        internal bool HasError { get; set; }
        internal string RequestedFileName { get; set; }
        internal bool NeedsBorderFile { get; set; }
        internal bool NeedsColorFile { get; set; }
        internal ExportFilePlan Plan { get; set; }

        internal void AddError(string text)
        {
            if (string.IsNullOrEmpty(text))
                return;
            HasError = true;
            Status = string.IsNullOrEmpty(Status) ? text : Status + "; " + text;
        }
    }

    internal sealed class FormExportNamePreview : System.Windows.Forms.Form
    {
        internal FormExportNamePreview(List<ExportNameRow> rows)
        {
            Text = "KPLN: Предпроверка имён файлов";
            StartPosition = FormStartPosition.CenterParent;
            Size = new Size(1120, 600);
            MinimumSize = new Size(800, 400);
            ShowInTaskbar = false;

            Label header = new Label
            {
                Dock = DockStyle.Top,
                Height = 58,
                Padding = new Padding(10),
                Text = "Текущие настройки и отмеченные листы/виды. Ошибок: " + rows.Count(r => r.HasError) +
                    ".\r\nЧтобы изменить настройки, закройте это окно и откройте проверку заново. Копирование строк: Ctrl+C."
            };
            DataGridView grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoGenerateColumns = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
                BackgroundColor = SystemColors.Window
            };
            grid.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            AddColumn(grid, "Format", "Формат", 60);
            AddColumn(grid, "Model", "Модель", 145);
            AddColumn(grid, "Sheet", "Лист / вид", 190);
            AddColumn(grid, "FileName", "Имя файла", 330);
            AddColumn(grid, "Status", "Результат проверки", 270);
            grid.DataSource = rows;
            grid.DataBindingComplete += (sender, args) =>
            {
                foreach (DataGridViewRow row in grid.Rows)
                    if (((ExportNameRow)row.DataBoundItem).HasError)
                        row.DefaultCellStyle.BackColor = System.Drawing.Color.MistyRose;
            };

            Button close = new Button { Text = "Закрыть", DialogResult = DialogResult.Cancel, AutoSize = true };
            FlowLayoutPanel footer = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom, Height = 42, Padding = new Padding(7),
                FlowDirection = FlowDirection.RightToLeft
            };
            footer.Controls.Add(close);
            CancelButton = close;
            Controls.Add(grid);
            Controls.Add(header);
            Controls.Add(footer);
        }

        private static void AddColumn(DataGridView grid, string property, string title, int width)
        {
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = property, HeaderText = title, Width = width,
                SortMode = DataGridViewColumnSortMode.NotSortable
            });
        }

        internal static List<ExportNameRow> BuildRows(string mainDocTitle,
            Dictionary<string, List<MainEntity>> selected, YayPrintSettings settings, bool printToFile)
        {
            List<ExportNameRow> rows = new List<ExportNameRow>();
            bool pdf = settings.isPDFExport && printToFile;
            bool merge = pdf && settings.isMergePdfs && settings.printerName == "PDFCreator";
            foreach (var model in selected)
            {
                ILookup<string, FamilyInstance> blocks = pdf && model.Value.Any(e => e.MainView is ViewSheet)
                    ? new FilteredElementCollector(model.Value[0].MainView.Document)
                        .OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType()
                        .Cast<FamilyInstance>()
                        .Where(t => t.get_Parameter(BuiltInParameter.SHEET_HEIGHT).AsDouble() > 0.6)
                        .ToLookup(t => t.get_Parameter(BuiltInParameter.SHEET_NUMBER).AsString())
                    : null;

                foreach (MainEntity entity in model.Value)
                {
                    if (pdf)
                    {
                        string name = FileNameTemplate.WithExtension(entity.NameByConstructor(settings.GetNameConstructor(entity.MainView), out List<string> errors), ".pdf");
                        int count = entity.MainView is ViewSheet sheet ? blocks[sheet.SheetNumber].Count() : 0;
                        for (int i = 0; i < Math.Max(1, count); i++)
                        {
                            string partName = count > 1 ? name.Replace(".pdf", "_" + i + ".pdf") : name;
                            ExportNameRow row = AddRow(rows, "PDF", model.Key, entity.ToString(), partName, errors);
                            row.NeedsBorderFile = settings.isExcludeBorders && entity.MainView is ViewSheet;
                            row.NeedsColorFile = settings.colorsType == ColorType.MonochromeWithExcludes && !entity.ForceColored;
                            if (entity.MainView is ViewSheet && count == 0)
                                row.AddError("Нет основной надписи для печати PDF");
                            if (!entity.MainView.CanBePrinted)
                                row.AddError("Вид недоступен для печати");
                            if (merge && !row.HasError)
                                row.Status = "Промежуточный PDF для объединения";
                        }
                    }
                    if (settings.isDWGExport)
                    {
                        string name = entity.NameByConstructor(settings.GetNameConstructor(entity.MainView), out List<string> errors);
                        name = FileNameTemplate.WithExtension(name, ".dwg");
                        AddRow(rows, "DWG", model.Key, entity.ToString(), name, errors);
                    }
                }
            }

            if (merge && selected.Count != 0)
                AddRow(rows, "PDF", mainDocTitle, "Объединённый файл", mainDocTitle + ".pdf", new List<string>());

            // Все выбранные модели пишут в одну папку соответствующего формата.
            foreach (var group in rows.GroupBy(r => r.Format + "|" + r.FileName, StringComparer.OrdinalIgnoreCase))
                if (group.Count() > 1)
                    foreach (ExportNameRow row in group)
                        row.AddError("Совпадение имён: " + group.Count() + " файла/-ов");

            foreach (var formatRows in rows.GroupBy(row => row.Format))
            {
                if (formatRows.Any(row => row.HasError))
                    continue;

                try
                {
                    List<string> names = new List<string>();
                    foreach (ExportNameRow row in formatRows)
                    {
                        names.Add(row.RequestedFileName);
                        if (row.NeedsBorderFile)
                            names.Add(FileNameTemplate.AppendSuffix(row.RequestedFileName, "_OUT_Border"));
                        if (row.NeedsColorFile)
                            names.Add(FileNameTemplate.AppendSuffix(row.RequestedFileName, "_OUT_Color"));
                    }

                    string root = formatRows.Key == "PDF" ? settings.outputPDFFolder : settings.outputDWGFolder;
                    ExportFilePlan plan = ExportFilePlan.Create(root, mainDocTitle, settings, formatRows.Key, names);
                    foreach (ExportNameRow row in formatRows)
                    {
                        row.Plan = plan;
                        row.FileName = plan.Names[row.RequestedFileName];
                        if (row.FileName != row.RequestedFileName)
                            row.Status = "Имя занято: добавлен порядковый номер";
                        if (plan.Redirected)
                            row.Status += (string.IsNullOrEmpty(row.Status) ? "" : "; ") +
                                "Длинный путь: выдача формата будет перенесена на системный диск";
                    }
                }
                catch (Exception ex)
                {
                    foreach (ExportNameRow row in formatRows)
                        row.AddError("Не удалось подготовить имена файлов: " + ex.Message);
                }
            }

            foreach (ExportNameRow row in rows)
            {
                if (string.IsNullOrEmpty(row.Status))
                    row.Status = "ОК";
            }

            if (rows.Count == 0)
                rows.Add(new ExportNameRow
                {
                    Status = "Отметьте листы/виды и включите PDF или DWG. Для бумажной печати имён файлов нет.",
                    HasError = selected.Count == 0 || (!settings.isPDFExport && !settings.isDWGExport)
                });
            return rows;
        }

        private static ExportNameRow AddRow(List<ExportNameRow> rows, string format, string model,
            string sheet, string name, List<string> errors)
        {
            ExportNameRow row = new ExportNameRow { Format = format, Model = model, Sheet = sheet, FileName = name, RequestedFileName = name };
            foreach (string error in errors)
                row.AddError(error);
            row.AddError(FileNameTemplate.Validate(name, format == "PDF" ? ".pdf" : ".dwg"));
            rows.Add(row);
            return row;
        }
    }
}
