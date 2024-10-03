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
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Publication.Common;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;


namespace KPLN_Publication
{
    public partial class FormPrint : System.Windows.Forms.Form
    {
        private readonly YayPrintSettings _printSettings;

        private readonly ToolTip _toolTip = new ToolTip();

        public YayPrintSettings PrintSettings
        {
            get { return _printSettings; }
        }

        private readonly Dictionary<string, List<MySheet>> _sheetsBaseToPrint = new Dictionary<string, List<MySheet>>();

        public Dictionary<string, List<MySheet>> _sheetsSelected = new Dictionary<string, List<MySheet>>();

        public bool printToFile = false;

        public FormPrint(Document doc, Dictionary<string, List<MySheet>> SheetsBase, YayPrintSettings printSettings, DBUser currentDBUser)
        {
            InitializeComponent();
            this.AcceptButton = btnOk;
            this.CancelButton = btnCancel;
            this.pluginVersion.Text = $"v.{ModuleData.Version}";

            int userDepartment = currentDBUser.SubDepartmentId;
            if (userDepartment == 2 || userDepartment == 8)
                this.checkBoxExcludeBorders.Enabled = true;

            _sheetsBaseToPrint = SheetsBase;

            //заполняю treeView
            foreach (var docWithSheets in _sheetsBaseToPrint)
            {
                TreeNode docNode = new TreeNode(docWithSheets.Key);
                bool haveChecked = false;
                foreach (MySheet sheet in docWithSheets.Value)
                {
                    string sheetTitle = sheet.ToString();
                    TreeNode sheetNode = new TreeNode(sheetTitle);
                    sheetNode.Checked = sheet.IsPrintable;
                    if (sheet.IsPrintable) haveChecked = true;
                    docNode.Nodes.Add(sheetNode);
                }
                if (haveChecked) docNode.Expand();
                treeView1.Nodes.Add(docNode);
            }

            #region Заполняю параметры печати PDF
            _printSettings = printSettings;
            checkBox_isPDFExport.Checked = printSettings.isPDFExport;
            textBox_PDFNameConstructor.Text = printSettings.pdfNameConstructor;
            textBox_PDFPath.Text = printSettings.outputPDFFolder;
            checkBoxMergePdfs.Checked = printSettings.isMergePdfs;
            checkBoxOrientation.Checked = printSettings.isUseOrientation;

            List<string> printers = new List<string>();
            foreach (string printer in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            {
                printers.Add(printer);
            }

            if (printers.Count == 0)
                throw new Exception("Cant find any installed printers");

            comboBoxPrinters.DataSource = printers;

            if (printers.Contains(_printSettings.printerName))
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

            if (_printSettings.hiddenLineProcessing == Autodesk.Revit.DB.HiddenLineViewsType.VectorProcessing)
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
            #endregion

            #region Заполняю параметры печати DWG
            checkBox_isDWGExport.Checked = printSettings.isDWGExport;
            textBox_DWGNameConstructor.Text = printSettings.dwgNameConstructor;
            textBox_DWGPath.Text = printSettings.outputDWGFolder;

            List<ExportDWGSettingsShell> dwgSettings = new FilteredElementCollector(doc)
                .OfClass(typeof(ExportDWGSettings))
                .Cast<ExportDWGSettings>()
                .Select(eds => new ExportDWGSettingsShell(eds.Name, eds))
                .ToList();
            comboBoxDWGExportTypes.DataSource = dwgSettings;
            comboBoxDWGExportTypes.DisplayMember = "Name";

            int dwgExpTypeFromConfigIndex = dwgSettings.FindIndex(ds => _printSettings.dwgExportSettingShell != null && _printSettings.dwgExportSettingShell.Name.Equals(ds.Name));
            if (dwgExpTypeFromConfigIndex != -1)
                comboBoxDWGExportTypes.SelectedIndex = dwgExpTypeFromConfigIndex;
            #endregion
        }

        private void cbx_Enter(object sender, EventArgs e)
        {
            string tt_text = "Выбери галку, чтобы скрыть рамки листа." +
                "\nВАЖНО:" +
                "\n1. Скрывается цвет RGD '003,002,051', если он используется " +
                "в проекте для документации - он тоже скроектся!" +
                "\n2. Работает ТОЛЬКО с принтером PDFCreator.";
            System.Windows.Forms.CheckBox cbx = sender as System.Windows.Forms.CheckBox;
            _toolTip.Show(tt_text, cbx);
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;

            foreach (TreeNode docNode in treeView1.Nodes)
            {
                string docNodeTitle = docNode.Text;
                //string revitDocTitle = sheetsBaseToPrint.Keys.Where(d => d == docNodeTitle).First();
                List<MySheet> selectedSheetsInDoc = new List<MySheet>();
                foreach (TreeNode sheetNode in docNode.Nodes)
                {
                    if (!sheetNode.Checked) continue;
                    string sheetTitle = sheetNode.Text;

                    var tempSheets = _sheetsBaseToPrint[docNodeTitle].Where(s => sheetTitle == s.ToString()).ToList();
                    if (tempSheets.Count == 0) throw new Exception("Cant get sheets from TreeNode");
                    MySheet msheet = tempSheets.First();
                    selectedSheetsInDoc.Add(msheet);
                }
                if (selectedSheetsInDoc.Count == 0) continue;

                _sheetsSelected.Add(docNodeTitle, selectedSheetsInDoc);
            }

            #region Обновление YayPrintSettings (изначально не реализован INotifyPrCh, продолжаю костыль)
            // Экспорт в PDF
            _printSettings.isPDFExport = checkBox_isPDFExport.Checked;

            _printSettings.printerName = comboBoxPrinters.SelectedItem.ToString();
            _printSettings.outputPDFFolder = textBox_PDFPath.Text;

            bool checkConstructor = false;
            string tempConstr = textBox_PDFNameConstructor.Text;
            if (tempConstr.Split('<').Length > 1)
            {
                if (tempConstr.Split('<')[1].Contains(">"))
                    checkConstructor = true;
            }
            if (checkConstructor)
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

            // Экспорт в DWG
            _printSettings.isDWGExport = checkBox_isDWGExport.Checked;
            _printSettings.dwgExportSettingShell = (ExportDWGSettingsShell)comboBoxDWGExportTypes.SelectedItem;
            _printSettings.outputDWGFolder = textBox_DWGPath.Text;
            _printSettings.dwgNameConstructor = textBox_DWGNameConstructor.Text;
            #endregion

            this.printToFile = radioButtonPDF.Checked;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void buttonPDFBrowse_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbDialog = new FolderBrowserDialog
            {
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
            formNameConstructor formName = new formNameConstructor(textBox_PDFNameConstructor.Text);

            if (formName.ShowDialog(this) == DialogResult.OK)
            {
                textBox_PDFNameConstructor.Text = formName.nameConstructor;
            }

            formName.Dispose();
        }

        private void btnDWGOpenNameConstructor_Click(object sender, EventArgs e)
        {
            formNameConstructor formName = new formNameConstructor(textBox_PDFNameConstructor.Text);

            if (formName.ShowDialog(this) == DialogResult.OK)
            {
                textBox_DWGNameConstructor.Text = formName.nameConstructor;
            }

            formName.Dispose();
        }

        private void radioButtonPDF_CheckedChanged(object sender, EventArgs e)
        {
            textBox_PDFPath.Enabled = true;
            buttonPDFBrowse.Enabled = true;
            textBox_PDFNameConstructor.Enabled = true;
            btnPDFOpenNameConstructor.Enabled = true;
            label5.Enabled = true;
            label6.Enabled = true;
            checkBoxMergePdfs.Enabled = true;
            checkBoxOrientation.Enabled = true;

            List<ColorType> colorTypes = Enum.GetValues(typeof(ColorType))
                .Cast<ColorType>()
                .ToList();
            comboBoxColors.DataSource = colorTypes;
            comboBoxColors.SelectedItem = _printSettings.colorsType;
        }

        private void radioButtonPaper_CheckedChanged(object sender, EventArgs e)
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

        private void checkBoxMergePdfs_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxMergePdfs.Checked == true)
            {
                textBox_PDFNameConstructor.Enabled = false;
                btnPDFOpenNameConstructor.Enabled = false;
                label6.Enabled = false;
            }
            else
            {
                textBox_PDFNameConstructor.Enabled = true;
                btnPDFOpenNameConstructor.Enabled = true;
                label6.Enabled = true;
            }
        }

        private void comboBoxColors_SelectedIndexChanged(object sender, EventArgs e)
        {
            ColorType curColorType = (ColorType)comboBoxColors.SelectedItem;
            if (curColorType == ColorType.MonochromeWithExcludes)
            {
                buttonExcludesColor.Enabled = true;
            }
            else
            {
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
            System.Diagnostics.Process.Start("http://moodle/mod/book/view.php?id=502&chapterid=667");
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
