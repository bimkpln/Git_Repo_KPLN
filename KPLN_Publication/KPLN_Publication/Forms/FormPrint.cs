#region License
/*Данный код опубликован под лицензией Creative Commons Attribution-ShareAlike.
Разрешено использовать, распространять, изменять и брать данный код за основу для производных в коммерческих и
некоммерческих целях, при условии указания авторства и если производные лицензируются на тех же условиях.
Код поставляется "как есть". Автор не несет ответственности за возможные последствия использования.
Зуев Александр, 2020, все права защищены.
This code is listed under the Creative Commons Attribution-ShareAlike license.
You may use, redistribute, remix, tweak, and build upon this work non-commercially and commercially,
as long as you credit the author by linking back and license your new creations under the same terms.
This code is provided 'as is'. Author disclaims any implied warranty.
Zuev Aleksandr, 2020, all rigths reserved.*/
#endregion

using Autodesk.Revit.DB;
using KPLN_Library_DBWorker;
using KPLN_Publication.Common;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;


namespace KPLN_Publication
{
    public partial class FormPrint : System.Windows.Forms.Form
    {
        private readonly YayPrintSettings _printSettings;
        private readonly string _mainDocTitle;
        private readonly bool _onlySheets;

        private readonly ToolTip _toolTip = new ToolTip();

        public YayPrintSettings PrintSettings
        {
            get => _printSettings;
        }

        private readonly Dictionary<string, List<MainEntity>> _entitiesBaseToPrint = new Dictionary<string, List<MainEntity>>();

        public Dictionary<string, List<MainEntity>> _entitiesSelected = new Dictionary<string, List<MainEntity>>();

        public bool printToFile = false;

        public FormPrint(Document doc, Dictionary<string, List<MainEntity>> mainEntites, YayPrintSettings printSettings)
        {
            InitializeComponent();
            _mainDocTitle = doc.Title;
            this.AcceptButton = btnOk;
            this.CancelButton = btnCancel;
            this.pluginVersion.Text = $"v.{ModuleData.Version}";

            int userDepartment = SQLiteMainService.CurrentDBUser.SubDepartmentId;
            if (userDepartment == 2 || userDepartment == 8)
                this.checkBoxExcludeBorders.Enabled = true;

            bool onlySheets = false;
            if (mainEntites.Values.FirstOrDefault().All(ent => ent.MainView is ViewSheet))
            {
                onlySheets = true;
                this.label4.Text = "Листы для выдачи";
            }
            else
                this.label4.Text = "Виды для выдачи";

            _onlySheets = onlySheets;
            _entitiesBaseToPrint = mainEntites;
            //заполняю treeView
            foreach (var docWithSheets in _entitiesBaseToPrint)
            {
                TreeNode docNode = new TreeNode(docWithSheets.Key);
                bool haveChecked = false;
                foreach (MainEntity sheet in docWithSheets.Value)
                {
                    string sheetTitle = sheet.ToString();
                    TreeNode sheetNode = new TreeNode(sheetTitle);
                    sheetNode.Checked = sheet.IsPrintable;
                    sheetNode.Tag = sheet;
                    if (sheet.IsPrintable) haveChecked = true;
                    docNode.Nodes.Add(sheetNode);
                }
                if (haveChecked) docNode.Expand();
                treeView1.Nodes.Add(docNode);
            }


            #region Заполняю параметры печати PDF
            _printSettings = printSettings;
            checkBox_isPDFExport.Checked = printSettings.isPDFExport;
            textBox_PDFNameConstructor.Text = FileNameTemplate.WithoutExtension(printSettings.pdfNameConstructor, ".pdf");
            textBox_PDFPath.Text = printSettings.outputPDFFolder;
            checkBoxMergePdfs.Checked = printSettings.isMergePdfs;
            checkBoxOrientation.Enabled = onlySheets;
            checkBoxOrientation.Checked = onlySheets && printSettings.isUseOrientation;
            checkBoxRefresh.Enabled = onlySheets;
            checkBoxRefresh.Checked = onlySheets && printSettings.isRefreshSchedules;
            checkBoxExcludeBorders.Enabled = onlySheets && checkBoxExcludeBorders.Enabled;
            checkBoxExcludeBorders.Checked = checkBoxExcludeBorders.Enabled && printSettings.isExcludeBorders;

            // Недоступное скрытие рамки сбрасываем и в конфиге, даже если форму закроют без печати.
            if (onlySheets && !checkBoxExcludeBorders.Enabled && _printSettings.isExcludeBorders)
            {
                checkBoxExcludeBorders.Checked = false;
                _printSettings.isExcludeBorders = false;
                YayPrintSettings.SaveSettings(_printSettings);
            }

            if (checkBox_isPDFExport.Enabled)
            {
                List<string> printers = new List<string>();
                foreach (string printer in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
                {
                    printers.Add(printer);
                }

                if (printers.Count == 0)
                    throw new Exception("Cant find any installed printers");

                comboBoxPrinters.DataSource = printers;

                if (_printSettings != null && printers.Contains(_printSettings.printerName))
                {
                    comboBoxPrinters.SelectedItem = printers.Where(i => i.Equals(_printSettings.printerName)).First();
                }
                else
                {
                    string selectedPrinterName = PrinterUtility.GetDefaultPrinter();
                    if (!printers.Contains(selectedPrinterName)) throw new Exception("Cant find printer " + selectedPrinterName);
                    comboBoxPrinters.SelectedItem = printers.Where(i => i.Equals(selectedPrinterName)).First();
                }

                radioButtonPDF.Checked = comboBoxPrinters.SelectedItem.ToString().Contains("PDF");
                radioButtonPaper.Checked = !radioButtonPDF.Checked;

                if (_printSettings != null && _printSettings.hiddenLineProcessing == Autodesk.Revit.DB.HiddenLineViewsType.VectorProcessing)
                {
                    radioButtonVector.Checked = true;
                    radioButtonRastr.Checked = false;
                }
                else
                {
                    radioButtonVector.Checked = false;
                    radioButtonRastr.Checked = true;
                }

                List<RasterQualityType> rasterTypes =
                    Enum.GetValues(typeof(RasterQualityType))
                    .Cast<RasterQualityType>()
                    .ToList();
                comboBoxRasterQuality.DataSource = rasterTypes;
                comboBoxRasterQuality.SelectedItem = _printSettings.rasterQuality;

                List<ColorType> colorTypes = Enum.GetValues(typeof(ColorType))
                    .Cast<ColorType>()
                    .ToList();
                comboBoxColors.DataSource = colorTypes;
                comboBoxColors.SelectedItem = _printSettings.colorsType;
            }
            #endregion

            #region Заполняю параметры печати DWG
            checkBox_isDWGExport.Checked = printSettings.isDWGExport;
            textBox_DWGNameConstructor.Text = FileNameTemplate.WithoutExtension(printSettings.dwgNameConstructor, ".dwg");
            textBox_DWGPath.Text = printSettings.outputDWGFolder;

            List<ExportDWGSettingsShell> dwgSettings = new FilteredElementCollector(doc)
                .OfClass(typeof(ExportDWGSettings))
                .Cast<ExportDWGSettings>()
                .Select(eds => new ExportDWGSettingsShell(eds.Name, eds))
                .ToList();
            comboBoxDWGExportTypes.DataSource = dwgSettings;
            comboBoxDWGExportTypes.DisplayMember = "Name";

            int dwgExpTypeFromConfigIndex = dwgSettings.FindIndex(ds =>
                _printSettings != null
                && _printSettings.dwgExportSettingShell != null
                && _printSettings.dwgExportSettingShell.Name.Equals(ds.Name));

            if (dwgExpTypeFromConfigIndex != -1)
                comboBoxDWGExportTypes.SelectedIndex = dwgExpTypeFromConfigIndex;
            #endregion

            checkBoxMergePdfs_CheckedChanged(null, null);

            textBox_PDFNameConstructor.TextChanged += NameRulesChanged;
            textBox_DWGNameConstructor.TextChanged += NameRulesChanged;
            checkBox_isPDFExport.CheckedChanged += NameRulesChanged;
            checkBox_isDWGExport.CheckedChanged += NameRulesChanged;
            radioButtonPDF.CheckedChanged += NameRulesChanged;
            radioButtonPaper.CheckedChanged += NameRulesChanged;
            treeView1.AfterCheck += (sender, args) => UpdateNameRulesState();
            UpdateNameRulesState();
        }

        private void cbx_Enter(object sender, EventArgs e)
        {
            string tt_text = "Выбери галку, чтобы скрыть рамки листа." +
                "\nВАЖНО:" +
                "\n1. Скрывается цвет RGB '003,002,051', если он используется " +
                "в проекте для документации - он тоже скроектся!" +
                "\n2. Работает ТОЛЬКО с принтером PDFCreator.";
            System.Windows.Forms.CheckBox cbx = sender as System.Windows.Forms.CheckBox;
            _toolTip.Show(tt_text, cbx);
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            UpdateNameRulesState();
            if (!btnOk.Enabled)
                return;

            _entitiesSelected = GetSelectedEntities();

            #region Обновление YayPrintSettings (изначально не реализован INotifyPrCh, продолжаю костыль)
            // Экспорт в PDF
            _printSettings.isPDFExport = checkBox_isPDFExport.Enabled && checkBox_isPDFExport.Checked;
            if (_printSettings.isPDFExport)
            {
                _printSettings.printerName = comboBoxPrinters.SelectedItem.ToString();
                _printSettings.outputPDFFolder = textBox_PDFPath.Text;

                _printSettings.pdfNameConstructor = textBox_PDFNameConstructor.Text;

                if (radioButtonVector.Checked)
                    _printSettings.hiddenLineProcessing = Autodesk.Revit.DB.HiddenLineViewsType.VectorProcessing;
                else
                    _printSettings.hiddenLineProcessing = Autodesk.Revit.DB.HiddenLineViewsType.RasterProcessing;

                _printSettings.rasterQuality = (Autodesk.Revit.DB.RasterQualityType)comboBoxRasterQuality.SelectedValue;

                _printSettings.isMergePdfs = checkBoxMergePdfs.Checked;
                _printSettings.isPrintToPaper = radioButtonPaper.Checked;

                _printSettings.colorsType = (ColorType)comboBoxColors.SelectedItem;
                _printSettings.isUseOrientation = checkBoxOrientation.Checked;
                _printSettings.isRefreshSchedules = checkBoxRefresh.Checked;
                _printSettings.isExcludeBorders = checkBoxExcludeBorders.Checked;
            }


            // Экспорт в DWG
            _printSettings.isDWGExport = checkBox_isDWGExport.Checked;
            if (_printSettings.isDWGExport)
            {
                _printSettings.dwgExportSettingShell = (ExportDWGSettingsShell)comboBoxDWGExportTypes.SelectedItem;
                _printSettings.outputDWGFolder = textBox_DWGPath.Text;
                _printSettings.dwgNameConstructor = textBox_DWGNameConstructor.Text;
            }
            #endregion

            this.printToFile = radioButtonPDF.Checked;
            List<ExportNameRow> rows = FormExportNamePreview.BuildRows(
                _mainDocTitle, _entitiesSelected, _printSettings, printToFile);
            if (rows.Any(row => row.HasError))
            {
                using (FormExportNamePreview preview = new FormExportNamePreview(rows))
                    preview.ShowDialog(this);
                return;
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void buttonPDFBrowse_Click(object sender, EventArgs e)
        {
            // Предустановка пути
            string presetPath = Path.GetPathRoot(Environment.SystemDirectory);
            if (new DirectoryInfo(this.textBox_PDFPath.Text).Exists)
                presetPath = this.textBox_PDFPath.Text;

            FolderBrowserDialog fbDialog = new FolderBrowserDialog
            {
                SelectedPath = presetPath,                
                ShowNewFolderButton = true
            };
            if (fbDialog.ShowDialog() == DialogResult.OK)
            {
                string path = fbDialog.SelectedPath;
                this.textBox_PDFPath.Text = path;
            }
        }

        private void buttonDWGBrowse_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbDialog = new FolderBrowserDialog
            {
                ShowNewFolderButton = true
            };
            if (fbDialog.ShowDialog() == DialogResult.OK)
            {
                string path = fbDialog.SelectedPath;
                textBox_DWGPath.Text = path;
            }
        }

        private void btnPDFOpenNameConstructor_Click(object sender, EventArgs e)
        {
            formNameConstructor formName = new formNameConstructor(textBox_PDFNameConstructor.Text, GetNameConstructorEntities());

            if (formName.ShowDialog(this) == DialogResult.OK)
            {
                textBox_PDFNameConstructor.Text = FileNameTemplate.WithoutExtension(formName.nameConstructor, ".pdf");
            }

            formName.Dispose();
        }

        private void btnDWGOpenNameConstructor_Click(object sender, EventArgs e)
        {
            formNameConstructor formName = new formNameConstructor(textBox_DWGNameConstructor.Text, GetNameConstructorEntities());

            if (formName.ShowDialog(this) == DialogResult.OK)
            {
                textBox_DWGNameConstructor.Text = FileNameTemplate.WithoutExtension(formName.nameConstructor, ".dwg");
            }

            formName.Dispose();
        }

        private void radioButtonPDF_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_isPDFExport.Enabled)
            {
                textBox_PDFPath.Enabled = true;
                buttonPDFBrowse.Enabled = true;
                textBox_PDFNameConstructor.Enabled = true;
                btnPDFOpenNameConstructor.Enabled = true;
                label5.Enabled = true;
                label6.Enabled = true;
                checkBoxMergePdfs.Enabled = true;
                checkBoxOrientation.Enabled = _onlySheets;

                List<ColorType> colorTypes = Enum.GetValues(typeof(ColorType))
                    .Cast<ColorType>()
                    .ToList();
                comboBoxColors.DataSource = colorTypes;
                comboBoxColors.SelectedItem = _printSettings.colorsType;
            }
        }

        private void radioButtonPaper_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_isPDFExport.Enabled)
            {
                textBox_PDFPath.Enabled = false;
                buttonPDFBrowse.Enabled = false;
                textBox_PDFNameConstructor.Enabled = false;
                btnPDFOpenNameConstructor.Enabled = false;
                label5.Enabled = false;
                label6.Enabled = false;
                checkBoxMergePdfs.Enabled = false;
                checkBoxOrientation.Enabled = false;

                List<ColorType> colorTypes = new List<ColorType> { ColorType.Color, ColorType.GrayScale, ColorType.Monochrome };
                comboBoxColors.DataSource = colorTypes;
                comboBoxColors.SelectedItem = ColorType.Monochrome;
            }
        }

        private void checkBoxMergePdfs_CheckedChanged(object sender, EventArgs e)
        {
            // Шаблон нужен и при объединении: по нему создаются промежуточные PDF.
            bool enabled = checkBox_isPDFExport.Checked && radioButtonPDF.Checked;
            textBox_PDFNameConstructor.Enabled = enabled;
            btnPDFOpenNameConstructor.Enabled = enabled;
            label6.Enabled = enabled;
        }

        private void comboBoxColors_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (checkBox_isPDFExport.Enabled)
            {
                ColorType curColorType = (ColorType)comboBoxColors.SelectedItem;
                if (curColorType == ColorType.MonochromeWithExcludes)
                    buttonExcludesColor.Enabled = true;
                else
                    buttonExcludesColor.Enabled = false;
            }
        }

        private void comboBoxPrinters_SelectedIndexChanged(object sender, EventArgs e)
        {
            radioButtonPDF.Checked = comboBoxPrinters.SelectedItem.ToString().Contains("PDF");
            radioButtonPaper.Checked = !radioButtonPDF.Checked;
        }

        private void buttonHelp_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("http://moodle/mod/book/view.php?id=502&chapterid=667") { UseShellExecute = true });
        }

        private Dictionary<string, List<MainEntity>> GetSelectedEntities()
        {
            Dictionary<string, List<MainEntity>> result = new Dictionary<string, List<MainEntity>>();
            foreach (TreeNode model in treeView1.Nodes)
            {
                List<MainEntity> selected = model.Nodes.Cast<TreeNode>()
                    .Where(node => node.Checked).Select(node => (MainEntity)node.Tag).ToList();
                if (selected.Count != 0)
                    result.Add(model.Text, selected);
            }
            return result;
        }

        private IEnumerable<MainEntity> GetNameConstructorEntities()
        {
            List<MainEntity> selected = GetSelectedEntities().Values.SelectMany(items => items).ToList();
            return selected.Count != 0 ? selected : _entitiesBaseToPrint.Values.SelectMany(items => items);
        }

        private void NameRulesChanged(object sender, EventArgs e)
        {
            checkBoxMergePdfs_CheckedChanged(sender, e);
            UpdateNameRulesState();
        }

        private void UpdateNameRulesState()
        {
            List<MainEntity> selected = GetSelectedEntities().Values.SelectMany(items => items).ToList();
            List<string> invalidFormats = GetInvalidNameFormats(selected,
                checkBox_isPDFExport.Checked && radioButtonPDF.Checked, textBox_PDFNameConstructor.Text,
                checkBox_isDWGExport.Checked, textBox_DWGNameConstructor.Text);

            labelNameRulesError.Text = invalidFormats.Count == 0 ? string.Empty
                : "Проверь правила построения имён для " + string.Join(" и ", invalidFormats);
            labelNameRulesError.Visible = invalidFormats.Count != 0;
            btnOk.Enabled = selected.Count != 0 && invalidFormats.Count == 0
                && (checkBox_isPDFExport.Checked || checkBox_isDWGExport.Checked);
        }

        internal static List<string> GetInvalidNameFormats(IEnumerable<MainEntity> selected,
            bool pdf, string pdfTemplate, bool dwg, string dwgTemplate)
        {
            List<string> result = new List<string>();
            foreach (string format in new[] { "PDF", "DWG" })
            {
                if ((format == "PDF" && !pdf) || (format == "DWG" && !dwg))
                    continue;

                foreach (MainEntity entity in selected)
                {
                    string name = entity.NameByConstructor(format == "PDF" ? pdfTemplate : dwgTemplate,
                        out List<string> errors);
                    name = FileNameTemplate.WithExtension(name, format == "PDF" ? ".pdf" : ".dwg");
                    if (errors.Count != 0 || FileNameTemplate.Validate(name, format == "PDF" ? ".pdf" : ".dwg") != string.Empty)
                    {
                        result.Add(format);
                        break;
                    }
                }
            }
            return result;
        }

        private void buttonPreview_Click(object sender, EventArgs e)
        {
            // Снимок текущих полей без сохранения настроек и изменения выбора.
            YayPrintSettings settings = new YayPrintSettings
            {
                isPDFExport = checkBox_isPDFExport.Enabled && checkBox_isPDFExport.Checked,
                isDWGExport = checkBox_isDWGExport.Checked,
                pdfNameConstructor = textBox_PDFNameConstructor.Text,
                dwgNameConstructor = textBox_DWGNameConstructor.Text,
                outputPDFFolder = textBox_PDFPath.Text,
                outputDWGFolder = textBox_DWGPath.Text,
                printerName = comboBoxPrinters.SelectedItem as string,
                isMergePdfs = checkBoxMergePdfs.Checked
            };
            List<ExportNameRow> rows = FormExportNamePreview.BuildRows(
                _mainDocTitle, GetSelectedEntities(), settings, radioButtonPDF.Checked);
            using (FormExportNamePreview preview = new FormExportNamePreview(rows))
                preview.ShowDialog(this);
        }

        private void radioButtonRastr_CheckedChanged(object sender, EventArgs e)
        {
            comboBoxRasterQuality.Enabled = true;
        }

        private void radioButtonVector_CheckedChanged(object sender, EventArgs e)
        {
            comboBoxRasterQuality.Enabled = false;
        }

        private void buttonExcludesColor_Click(object sender, EventArgs e)
        {
            FormExcludeColors form = new FormExcludeColors(_printSettings.excludeColors);
            if (form.ShowDialog() != DialogResult.OK) return;

            PrintSettings.excludeColors = form.Colors;
        }
    }
}
