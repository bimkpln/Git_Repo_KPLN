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
        private readonly bool _canExcludeBorders;
        private readonly int _nameRulesFullHeight;
        private readonly int _nameRulesRowHeight;

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
            _nameRulesFullHeight = groupBoxNameRules.Height;
            _nameRulesRowHeight = labelViewName.Top - labelSheetName.Top;
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
                this.label4.Text = mainEntites.Values.SelectMany(items => items).Any(ent => ent.MainView is ViewSheet)
                    ? "Листы и виды для выдачи" : "Виды для выдачи";

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
            _printSettings.InitializeNameConstructors();
            checkBox_isPDFExport.Checked = printSettings.isPDFExport;
            textBox_SheetNameConstructor.Text = printSettings.sheetNameConstructor;
            textBox_ViewNameConstructor.Text = printSettings.viewNameConstructor;
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

            _canExcludeBorders = checkBoxExcludeBorders.Enabled;
            UpdateExportControlsState();
            checkBoxMergePdfs_CheckedChanged(null, null);

            textBox_SheetNameConstructor.TextChanged += NameRulesChanged;
            textBox_ViewNameConstructor.TextChanged += NameRulesChanged;
            checkBox_isPDFExport.CheckedChanged += NameRulesChanged;
            checkBox_isDWGExport.CheckedChanged += NameRulesChanged;
            radioButtonPDF.CheckedChanged += NameRulesChanged;
            radioButtonPaper.CheckedChanged += NameRulesChanged;
            textBox_PDFPath.TextChanged += NameRulesChanged;
            textBox_DWGPath.TextChanged += NameRulesChanged;
            checkBoxMergePdfs.CheckedChanged += NameRulesChanged;
            checkBoxExcludeBorders.CheckedChanged += NameRulesChanged;
            comboBoxColors.SelectedIndexChanged += NameRulesChanged;
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
            _printSettings.sheetNameConstructor = textBox_SheetNameConstructor.Text;
            _printSettings.viewNameConstructor = textBox_ViewNameConstructor.Text;

            #region Обновление YayPrintSettings (изначально не реализован INotifyPrCh, продолжаю костыль)
            // Экспорт в PDF
            _printSettings.isPDFExport = checkBox_isPDFExport.Enabled && checkBox_isPDFExport.Checked;
            if (_printSettings.isPDFExport)
            {
                _printSettings.printerName = comboBoxPrinters.SelectedItem.ToString();
                _printSettings.outputPDFFolder = textBox_PDFPath.Text;


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
            }
            #endregion

            this.printToFile = radioButtonPDF.Checked;
            List<ExportNameRow> rows = FormExportNamePreview.BuildRows(
                _mainDocTitle, _entitiesSelected, _printSettings, printToFile);
            if (rows.Any(row => row.HasError))
            {
                ApplyNameRulesState(rows);
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

        private void btnSheetOpenNameConstructor_Click(object sender, EventArgs e)
        {
            formNameConstructor formName = new formNameConstructor(textBox_SheetNameConstructor.Text, GetNameConstructorEntities(true));
            formName.Text = "Конструктор имени — листы";

            if (formName.ShowDialog(this) == DialogResult.OK)
            {
                textBox_SheetNameConstructor.Text = FileNameTemplate.WithoutExportExtension(formName.nameConstructor);
            }

            formName.Dispose();
        }

        private void btnViewOpenNameConstructor_Click(object sender, EventArgs e)
        {
            formNameConstructor formName = new formNameConstructor(textBox_ViewNameConstructor.Text, GetNameConstructorEntities(false));
            formName.Text = "Конструктор имени — виды";

            if (formName.ShowDialog(this) == DialogResult.OK)
            {
                textBox_ViewNameConstructor.Text = FileNameTemplate.WithoutExportExtension(formName.nameConstructor);
            }

            formName.Dispose();
        }

        private void radioButtonPDF_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_isPDFExport.Enabled)
            {
                textBox_PDFPath.Enabled = true;
                buttonPDFBrowse.Enabled = true;
                label5.Enabled = true;
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
                label5.Enabled = false;
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
            groupBoxNameRules.Enabled = (checkBox_isPDFExport.Checked && radioButtonPDF.Checked)
                || checkBox_isDWGExport.Checked;
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

        private IEnumerable<MainEntity> GetNameConstructorEntities(bool sheets)
        {
            List<MainEntity> selected = GetSelectedEntities().Values.SelectMany(items => items)
                .Where(entity => (entity.MainView is ViewSheet) == sheets).ToList();
            return selected.Count != 0 ? selected : _entitiesBaseToPrint.Values.SelectMany(items => items)
                .Where(entity => (entity.MainView is ViewSheet) == sheets);
        }

        private void UpdateNameConstructorLayout(bool sheets, bool views)
        {
            labelSheetName.Visible = textBox_SheetNameConstructor.Visible = btnSheetOpenNameConstructor.Visible = sheets;
            labelViewName.Visible = textBox_ViewNameConstructor.Visible = btnViewOpenNameConstructor.Visible = views;

            int offset = sheets ? _nameRulesRowHeight : 0;
            labelViewName.Top = labelSheetName.Top + offset;
            textBox_ViewNameConstructor.Top = textBox_SheetNameConstructor.Top + offset;
            btnViewOpenNameConstructor.Top = btnSheetOpenNameConstructor.Top + offset;

            int height = _nameRulesFullHeight - (sheets && views ? 0 : _nameRulesRowHeight);
            int delta = height - groupBoxNameRules.Height;
            if (delta != 0)
            {
                groupBoxNameRules.Height = height;
                Height += delta;
            }
        }

        private void UpdateExportControlsState()
        {
            bool printEnabled = checkBox_isPDFExport.Enabled && checkBox_isPDFExport.Checked;
            bool pdfEnabled = printEnabled && radioButtonPDF.Checked;
            bool dwgEnabled = checkBox_isDWGExport.Enabled && checkBox_isDWGExport.Checked;

            // Переключатели остаются доступны, значения отключённых полей сохраняются.
            foreach (System.Windows.Forms.Control control in groupBox1.Controls)
            {
                if (control != checkBox_isPDFExport)
                    control.Enabled = printEnabled;
            }

            groupBox2.Enabled = printEnabled;
            groupBox3.Enabled = printEnabled;
            textBox_PDFPath.Enabled = buttonPDFBrowse.Enabled = label5.Enabled = pdfEnabled;
            checkBoxMergePdfs.Enabled = pdfEnabled;
            checkBoxOrientation.Enabled = pdfEnabled && _onlySheets;
            checkBoxRefresh.Enabled = printEnabled && _onlySheets;
            checkBoxExcludeBorders.Enabled = printEnabled && _canExcludeBorders;
            comboBoxRasterQuality.Enabled = printEnabled && radioButtonRastr.Checked;
            buttonExcludesColor.Enabled = printEnabled
                && comboBoxColors.SelectedItem is ColorType color
                && color == ColorType.MonochromeWithExcludes;

            foreach (System.Windows.Forms.Control control in groupBox4.Controls)
            {
                if (control != checkBox_isDWGExport)
                    control.Enabled = dwgEnabled;
            }

            groupBoxNameRules.Enabled = pdfEnabled || dwgEnabled;
        }

        private void NameRulesChanged(object sender, EventArgs e)
        {
            UpdateExportControlsState();
            checkBoxMergePdfs_CheckedChanged(sender, e);
            UpdateNameRulesState();
        }

        private List<ExportNameRow> UpdateNameRulesState()
        {
            List<MainEntity> selected = GetSelectedEntities().Values.SelectMany(items => items).ToList();
            IEnumerable<MainEntity> visible = selected.Count != 0 ? selected : _entitiesBaseToPrint.Values.SelectMany(items => items);
            UpdateNameConstructorLayout(visible.Any(entity => entity.MainView is ViewSheet),
                visible.Any(entity => !(entity.MainView is ViewSheet)));
            List<ExportNameRow> rows = BuildCurrentNameRows();
            ApplyNameRulesState(rows);
            return rows;
        }

        internal static bool HasNameErrors(IEnumerable<ExportNameRow> rows)
        {
            return rows.Any(row => row.HasError);
        }

        private void ApplyNameRulesState(List<ExportNameRow> rows)
        {
            bool hasErrors = HasNameErrors(rows);
            List<string> invalidFormats = rows.Where(row => row.HasError && !string.IsNullOrEmpty(row.Format))
                .Select(row => row.Format).Distinct().ToList();

            labelNameRulesError.Text = !hasErrors ? string.Empty : invalidFormats.Count != 0
                ? "Проверь правила построения имён для " + string.Join(" и ", invalidFormats)
                : "Отметьте листы/виды и включите экспорт";
            labelNameRulesError.Visible = hasErrors;
            btnOk.Enabled = !hasErrors && GetSelectedEntities().Count != 0
                && (checkBox_isPDFExport.Checked || checkBox_isDWGExport.Checked);
        }

        private List<ExportNameRow> BuildCurrentNameRows()
        {
            // Снимок текущих полей без сохранения настроек и изменения выбора.
            YayPrintSettings settings = new YayPrintSettings
            {
                isPDFExport = checkBox_isPDFExport.Enabled && checkBox_isPDFExport.Checked,
                isDWGExport = checkBox_isDWGExport.Checked,
                sheetNameConstructor = textBox_SheetNameConstructor.Text,
                viewNameConstructor = textBox_ViewNameConstructor.Text,
                outputPDFFolder = textBox_PDFPath.Text,
                outputDWGFolder = textBox_DWGPath.Text,
                printerName = comboBoxPrinters.SelectedItem as string,
                isMergePdfs = checkBoxMergePdfs.Checked,
                isExcludeBorders = checkBoxExcludeBorders.Checked,
                colorsType = comboBoxColors.SelectedItem is ColorType color ? color : ColorType.Monochrome,
                ExportRunFolderName = _printSettings.GetExportFolderName(_mainDocTitle),
                ExportFallbackFolderName = _printSettings.GetExportFallbackFolderName()
            };
            return FormExportNamePreview.BuildRows(
                _mainDocTitle, GetSelectedEntities(), settings, radioButtonPDF.Checked);
        }

        private void buttonPreview_Click(object sender, EventArgs e)
        {
            List<ExportNameRow> rows = UpdateNameRulesState();
            using (FormExportNamePreview preview = new FormExportNamePreview(rows))
                preview.ShowDialog(this);
            UpdateNameRulesState();
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
