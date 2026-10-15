namespace KPLN_Publication
{
    partial class FormPrint
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.checkBoxExcludeBorders = new System.Windows.Forms.CheckBox();
            this.checkBoxRefresh = new System.Windows.Forms.CheckBox();
            this.checkBoxOrientation = new System.Windows.Forms.CheckBox();
            this.checkBox_isPDFExport = new System.Windows.Forms.CheckBox();
            this.checkBoxMergePdfs = new System.Windows.Forms.CheckBox();
            this.radioButtonPDF = new System.Windows.Forms.RadioButton();
            this.radioButtonPaper = new System.Windows.Forms.RadioButton();
            this.buttonPDFBrowse = new System.Windows.Forms.Button();
            this.textBox_PDFPath = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.comboBoxPrinters = new System.Windows.Forms.ComboBox();
            this.groupBoxNameRules = new System.Windows.Forms.GroupBox();
            this.btnSheetOpenNameConstructor = new System.Windows.Forms.Button();
            this.textBox_SheetNameConstructor = new System.Windows.Forms.TextBox();
            this.labelSheetName = new System.Windows.Forms.Label();
            this.labelViewName = new System.Windows.Forms.Label();
            this.textBox_ViewNameConstructor = new System.Windows.Forms.TextBox();
            this.btnViewOpenNameConstructor = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.radioButtonRastr = new System.Windows.Forms.RadioButton();
            this.radioButtonVector = new System.Windows.Forms.RadioButton();
            this.comboBoxRasterQuality = new System.Windows.Forms.ComboBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.buttonExcludesColor = new System.Windows.Forms.Button();
            this.comboBoxColors = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.pluginVersion = new System.Windows.Forms.Label();
            this.labelNameRulesError = new System.Windows.Forms.Label();
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.buttonPreview = new System.Windows.Forms.Button();
            this.buttonHelp = new System.Windows.Forms.Button();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.comboBoxDWGExportTypes = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.checkBox_isDWGExport = new System.Windows.Forms.CheckBox();
            this.label9 = new System.Windows.Forms.Label();
            this.textBox_DWGPath = new System.Windows.Forms.TextBox();
            this.buttonDWGBrowse = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBoxNameRules.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.checkBoxExcludeBorders);
            this.groupBox1.Controls.Add(this.checkBoxRefresh);
            this.groupBox1.Controls.Add(this.checkBoxOrientation);
            this.groupBox1.Controls.Add(this.checkBox_isPDFExport);
            this.groupBox1.Controls.Add(this.checkBoxMergePdfs);
            this.groupBox1.Controls.Add(this.radioButtonPDF);
            this.groupBox1.Controls.Add(this.radioButtonPaper);
            this.groupBox1.Controls.Add(this.buttonPDFBrowse);
            this.groupBox1.Controls.Add(this.textBox_PDFPath);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.comboBoxPrinters);
            this.groupBox1.Location = new System.Drawing.Point(250, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(277, 225);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Экспорт в PDF/печать";
            // 
            // checkBoxExcludeBorders
            // 
            this.checkBoxExcludeBorders.AutoSize = true;
            this.checkBoxExcludeBorders.Enabled = false;
            this.checkBoxExcludeBorders.Location = new System.Drawing.Point(9, 204);
            this.checkBoxExcludeBorders.Margin = new System.Windows.Forms.Padding(2);
            this.checkBoxExcludeBorders.Name = "checkBoxExcludeBorders";
            this.checkBoxExcludeBorders.Size = new System.Drawing.Size(250, 17);
            this.checkBoxExcludeBorders.TabIndex = 14;
            this.checkBoxExcludeBorders.Text = "Исключить границы листа (только АР_КОН)";
            this.checkBoxExcludeBorders.UseVisualStyleBackColor = true;
            this.checkBoxExcludeBorders.MouseEnter += new System.EventHandler(this.cbx_Enter);
            // 
            // checkBoxRefresh
            // 
            this.checkBoxRefresh.AutoSize = true;
            this.checkBoxRefresh.Checked = true;
            this.checkBoxRefresh.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxRefresh.Location = new System.Drawing.Point(9, 183);
            this.checkBoxRefresh.Margin = new System.Windows.Forms.Padding(2);
            this.checkBoxRefresh.Name = "checkBoxRefresh";
            this.checkBoxRefresh.Size = new System.Drawing.Size(152, 17);
            this.checkBoxRefresh.TabIndex = 13;
            this.checkBoxRefresh.Text = "Обновить спецификации";
            this.checkBoxRefresh.UseVisualStyleBackColor = true;
            // 
            // checkBoxOrientation
            // 
            this.checkBoxOrientation.AutoSize = true;
            this.checkBoxOrientation.Location = new System.Drawing.Point(9, 162);
            this.checkBoxOrientation.Margin = new System.Windows.Forms.Padding(2);
            this.checkBoxOrientation.Name = "checkBoxOrientation";
            this.checkBoxOrientation.Size = new System.Drawing.Size(219, 17);
            this.checkBoxOrientation.TabIndex = 13;
            this.checkBoxOrientation.Text = "Улучшенное определение ориентации";
            this.checkBoxOrientation.UseVisualStyleBackColor = true;
            // 
            // checkBox_isPDFExport
            // 
            this.checkBox_isPDFExport.AutoSize = true;
            this.checkBox_isPDFExport.Location = new System.Drawing.Point(9, 18);
            this.checkBox_isPDFExport.Margin = new System.Windows.Forms.Padding(2);
            this.checkBox_isPDFExport.Name = "checkBox_isPDFExport";
            this.checkBox_isPDFExport.Size = new System.Drawing.Size(185, 17);
            this.checkBox_isPDFExport.TabIndex = 13;
            this.checkBox_isPDFExport.Text = "Выполнить экспорт на бумагу?";
            this.checkBox_isPDFExport.UseVisualStyleBackColor = true;
            this.checkBox_isPDFExport.CheckedChanged += new System.EventHandler(this.checkBoxMergePdfs_CheckedChanged);
            // 
            // checkBoxMergePdfs
            // 
            this.checkBoxMergePdfs.AutoSize = true;
            this.checkBoxMergePdfs.Location = new System.Drawing.Point(9, 141);
            this.checkBoxMergePdfs.Margin = new System.Windows.Forms.Padding(2);
            this.checkBoxMergePdfs.Name = "checkBoxMergePdfs";
            this.checkBoxMergePdfs.Size = new System.Drawing.Size(153, 17);
            this.checkBoxMergePdfs.TabIndex = 13;
            this.checkBoxMergePdfs.Text = "Объединить в один файл";
            this.checkBoxMergePdfs.UseVisualStyleBackColor = true;
            this.checkBoxMergePdfs.CheckedChanged += new System.EventHandler(this.checkBoxMergePdfs_CheckedChanged);
            // 
            // radioButtonPDF
            // 
            this.radioButtonPDF.AutoSize = true;
            this.radioButtonPDF.Checked = true;
            this.radioButtonPDF.Location = new System.Drawing.Point(122, 80);
            this.radioButtonPDF.Name = "radioButtonPDF";
            this.radioButtonPDF.Size = new System.Drawing.Size(56, 17);
            this.radioButtonPDF.TabIndex = 12;
            this.radioButtonPDF.TabStop = true;
            this.radioButtonPDF.Text = "В PDF";
            this.radioButtonPDF.UseVisualStyleBackColor = true;
            this.radioButtonPDF.CheckedChanged += new System.EventHandler(this.radioButtonPDF_CheckedChanged);
            // 
            // radioButtonPaper
            // 
            this.radioButtonPaper.AutoSize = true;
            this.radioButtonPaper.Location = new System.Drawing.Point(9, 80);
            this.radioButtonPaper.Name = "radioButtonPaper";
            this.radioButtonPaper.Size = new System.Drawing.Size(77, 17);
            this.radioButtonPaper.TabIndex = 11;
            this.radioButtonPaper.Text = "На бумагу";
            this.radioButtonPaper.UseVisualStyleBackColor = true;
            this.radioButtonPaper.CheckedChanged += new System.EventHandler(this.radioButtonPaper_CheckedChanged);
            // 
            // buttonPDFBrowse
            // 
            this.buttonPDFBrowse.Enabled = false;
            this.buttonPDFBrowse.Location = new System.Drawing.Point(238, 116);
            this.buttonPDFBrowse.Name = "buttonPDFBrowse";
            this.buttonPDFBrowse.Size = new System.Drawing.Size(27, 20);
            this.buttonPDFBrowse.TabIndex = 6;
            this.buttonPDFBrowse.Text = "...";
            this.buttonPDFBrowse.UseVisualStyleBackColor = true;
            this.buttonPDFBrowse.Click += new System.EventHandler(this.buttonPDFBrowse_Click);
            // 
            // textBox_PDFPath
            // 
            this.textBox_PDFPath.Location = new System.Drawing.Point(9, 116);
            this.textBox_PDFPath.Name = "textBox_PDFPath";
            this.textBox_PDFPath.Size = new System.Drawing.Size(223, 20);
            this.textBox_PDFPath.TabIndex = 5;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Enabled = false;
            this.label5.Location = new System.Drawing.Point(6, 100);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(117, 13);
            this.label5.TabIndex = 4;
            this.label5.Text = "Путь для сохранения:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 35);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(177, 13);
            this.label3.TabIndex = 3;
            this.label3.Text = "Имя (рекомендуется PDFCreator):";
            // 
            // comboBoxPrinters
            // 
            this.comboBoxPrinters.FormattingEnabled = true;
            this.comboBoxPrinters.Location = new System.Drawing.Point(9, 53);
            this.comboBoxPrinters.Margin = new System.Windows.Forms.Padding(8, 3, 8, 3);
            this.comboBoxPrinters.Name = "comboBoxPrinters";
            this.comboBoxPrinters.Size = new System.Drawing.Size(256, 21);
            this.comboBoxPrinters.TabIndex = 2;
            this.comboBoxPrinters.Text = "PDFCreator";
            this.comboBoxPrinters.SelectedIndexChanged += new System.EventHandler(this.comboBoxPrinters_SelectedIndexChanged);
            // 
            // groupBoxNameRules
            // 
            this.groupBoxNameRules.Controls.Add(this.btnSheetOpenNameConstructor);
            this.groupBoxNameRules.Controls.Add(this.textBox_SheetNameConstructor);
            this.groupBoxNameRules.Controls.Add(this.labelSheetName);
            this.groupBoxNameRules.Controls.Add(this.labelViewName);
            this.groupBoxNameRules.Controls.Add(this.textBox_ViewNameConstructor);
            this.groupBoxNameRules.Controls.Add(this.btnViewOpenNameConstructor);
            this.groupBoxNameRules.Location = new System.Drawing.Point(250, 496);
            this.groupBoxNameRules.Name = "groupBoxNameRules";
            this.groupBoxNameRules.Size = new System.Drawing.Size(278, 112);
            this.groupBoxNameRules.TabIndex = 3;
            this.groupBoxNameRules.TabStop = false;
            this.groupBoxNameRules.Text = "Имена файлов для PDF и DWG";
            // 
            // btnSheetOpenNameConstructor
            // 
            this.btnSheetOpenNameConstructor.Location = new System.Drawing.Point(238, 34);
            this.btnSheetOpenNameConstructor.Name = "btnSheetOpenNameConstructor";
            this.btnSheetOpenNameConstructor.Size = new System.Drawing.Size(27, 21);
            this.btnSheetOpenNameConstructor.TabIndex = 9;
            this.btnSheetOpenNameConstructor.Text = "...";
            this.btnSheetOpenNameConstructor.UseVisualStyleBackColor = true;
            this.btnSheetOpenNameConstructor.Click += new System.EventHandler(this.btnSheetOpenNameConstructor_Click);
            // 
            // textBox_SheetNameConstructor
            // 
            this.textBox_SheetNameConstructor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_SheetNameConstructor.Location = new System.Drawing.Point(9, 34);
            this.textBox_SheetNameConstructor.Name = "textBox_SheetNameConstructor";
            this.textBox_SheetNameConstructor.ReadOnly = true;
            this.textBox_SheetNameConstructor.Size = new System.Drawing.Size(223, 20);
            this.textBox_SheetNameConstructor.TabIndex = 8;
            // 
            // labelSheetName
            // 
            this.labelSheetName.AutoSize = true;
            this.labelSheetName.Location = new System.Drawing.Point(6, 18);
            this.labelSheetName.Name = "labelSheetName";
            this.labelSheetName.Size = new System.Drawing.Size(43, 13);
            this.labelSheetName.TabIndex = 7;
            this.labelSheetName.Text = "Листы:";
            // 
            // labelViewName
            // 
            this.labelViewName.AutoSize = true;
            this.labelViewName.Location = new System.Drawing.Point(6, 61);
            this.labelViewName.Name = "labelViewName";
            this.labelViewName.Size = new System.Drawing.Size(37, 13);
            this.labelViewName.TabIndex = 7;
            this.labelViewName.Text = "Виды:";
            // 
            // textBox_ViewNameConstructor
            // 
            this.textBox_ViewNameConstructor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_ViewNameConstructor.Location = new System.Drawing.Point(9, 77);
            this.textBox_ViewNameConstructor.Name = "textBox_ViewNameConstructor";
            this.textBox_ViewNameConstructor.ReadOnly = true;
            this.textBox_ViewNameConstructor.Size = new System.Drawing.Size(223, 20);
            this.textBox_ViewNameConstructor.TabIndex = 10;
            // 
            // btnViewOpenNameConstructor
            // 
            this.btnViewOpenNameConstructor.Location = new System.Drawing.Point(238, 77);
            this.btnViewOpenNameConstructor.Name = "btnViewOpenNameConstructor";
            this.btnViewOpenNameConstructor.Size = new System.Drawing.Size(27, 21);
            this.btnViewOpenNameConstructor.TabIndex = 11;
            this.btnViewOpenNameConstructor.Text = "...";
            this.btnViewOpenNameConstructor.UseVisualStyleBackColor = true;
            this.btnViewOpenNameConstructor.Click += new System.EventHandler(this.btnViewOpenNameConstructor_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.radioButtonRastr);
            this.groupBox2.Controls.Add(this.radioButtonVector);
            this.groupBox2.Controls.Add(this.comboBoxRasterQuality);
            this.groupBox2.Location = new System.Drawing.Point(250, 243);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(278, 50);
            this.groupBox2.TabIndex = 0;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Обработка";
            // 
            // radioButtonRastr
            // 
            this.radioButtonRastr.AutoSize = true;
            this.radioButtonRastr.Location = new System.Drawing.Point(91, 19);
            this.radioButtonRastr.Name = "radioButtonRastr";
            this.radioButtonRastr.Size = new System.Drawing.Size(79, 17);
            this.radioButtonRastr.TabIndex = 2;
            this.radioButtonRastr.Text = "Растровая";
            this.radioButtonRastr.UseVisualStyleBackColor = true;
            this.radioButtonRastr.CheckedChanged += new System.EventHandler(this.radioButtonRastr_CheckedChanged);
            // 
            // radioButtonVector
            // 
            this.radioButtonVector.AutoSize = true;
            this.radioButtonVector.Checked = true;
            this.radioButtonVector.Location = new System.Drawing.Point(9, 19);
            this.radioButtonVector.Name = "radioButtonVector";
            this.radioButtonVector.Size = new System.Drawing.Size(79, 17);
            this.radioButtonVector.TabIndex = 1;
            this.radioButtonVector.TabStop = true;
            this.radioButtonVector.Text = "Векторная";
            this.radioButtonVector.UseVisualStyleBackColor = true;
            this.radioButtonVector.CheckedChanged += new System.EventHandler(this.radioButtonVector_CheckedChanged);
            // 
            // comboBoxRasterQuality
            // 
            this.comboBoxRasterQuality.Enabled = false;
            this.comboBoxRasterQuality.FormattingEnabled = true;
            this.comboBoxRasterQuality.Items.AddRange(new object[] {
            "Низкое",
            "Среднее",
            "Высокое",
            "Презентационное"});
            this.comboBoxRasterQuality.Location = new System.Drawing.Point(177, 18);
            this.comboBoxRasterQuality.Margin = new System.Windows.Forms.Padding(10, 3, 10, 3);
            this.comboBoxRasterQuality.Name = "comboBoxRasterQuality";
            this.comboBoxRasterQuality.Size = new System.Drawing.Size(88, 21);
            this.comboBoxRasterQuality.TabIndex = 1;
            this.comboBoxRasterQuality.Text = "Презентационное";
            // 
            // groupBox3
            // 
            this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox3.Controls.Add(this.buttonExcludesColor);
            this.groupBox3.Controls.Add(this.comboBoxColors);
            this.groupBox3.Controls.Add(this.label2);
            this.groupBox3.Location = new System.Drawing.Point(250, 299);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(278, 72);
            this.groupBox3.TabIndex = 1;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Вывод на печать";
            // 
            // buttonExcludesColor
            // 
            this.buttonExcludesColor.Location = new System.Drawing.Point(80, 41);
            this.buttonExcludesColor.Name = "buttonExcludesColor";
            this.buttonExcludesColor.Size = new System.Drawing.Size(192, 23);
            this.buttonExcludesColor.TabIndex = 3;
            this.buttonExcludesColor.Text = "Исключения цветов";
            this.buttonExcludesColor.UseVisualStyleBackColor = true;
            this.buttonExcludesColor.Click += new System.EventHandler(this.buttonExcludesColor_Click);
            // 
            // comboBoxColors
            // 
            this.comboBoxColors.FormattingEnabled = true;
            this.comboBoxColors.Items.AddRange(new object[] {
            "Черные линии",
            "Оттенки серого",
            "Цвет"});
            this.comboBoxColors.Location = new System.Drawing.Point(80, 15);
            this.comboBoxColors.Name = "comboBoxColors";
            this.comboBoxColors.Size = new System.Drawing.Size(192, 21);
            this.comboBoxColors.TabIndex = 2;
            this.comboBoxColors.Text = "Оттенки серого";
            this.comboBoxColors.SelectedIndexChanged += new System.EventHandler(this.comboBoxColors_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 18);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Цвета:";
            // 
            // btnOk
            // 
            this.btnOk.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOk.Location = new System.Drawing.Point(174, 620);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(75, 23);
            this.btnOk.TabIndex = 4;
            this.btnOk.Text = "ОК";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.Location = new System.Drawing.Point(255, 620);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 5;
            this.btnCancel.Text = "Отмена";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 12);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(104, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Листы для выдачи:";
            // 
            // pluginVersion
            // 
            this.pluginVersion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.pluginVersion.AutoSize = true;
            this.pluginVersion.Location = new System.Drawing.Point(12, 630);
            this.pluginVersion.Name = "pluginVersion";
            this.pluginVersion.Size = new System.Drawing.Size(67, 13);
            this.pluginVersion.TabIndex = 0;
            this.pluginVersion.Text = "v2020.09.09";
            // 
            // labelNameRulesError
            // 
            this.labelNameRulesError.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.labelNameRulesError.ForeColor = System.Drawing.Color.Red;
            this.labelNameRulesError.Location = new System.Drawing.Point(12, 654);
            this.labelNameRulesError.Name = "labelNameRulesError";
            this.labelNameRulesError.Size = new System.Drawing.Size(515, 20);
            this.labelNameRulesError.TabIndex = 4;
            this.labelNameRulesError.Visible = false;
            // 
            // treeView1
            // 
            this.treeView1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.treeView1.CheckBoxes = true;
            this.treeView1.FullRowSelect = true;
            this.treeView1.Location = new System.Drawing.Point(12, 28);
            this.treeView1.Name = "treeView1";
            this.treeView1.Size = new System.Drawing.Size(232, 580);
            this.treeView1.TabIndex = 9;
            // 
            // buttonPreview
            // 
            this.buttonPreview.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonPreview.Location = new System.Drawing.Point(336, 620);
            this.buttonPreview.Name = "buttonPreview";
            this.buttonPreview.Size = new System.Drawing.Size(110, 23);
            this.buttonPreview.TabIndex = 5;
            this.buttonPreview.Text = "Имена файлов";
            this.buttonPreview.UseVisualStyleBackColor = true;
            this.buttonPreview.Click += new System.EventHandler(this.buttonPreview_Click);
            // 
            // buttonHelp
            // 
            this.buttonHelp.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonHelp.Location = new System.Drawing.Point(452, 620);
            this.buttonHelp.Name = "buttonHelp";
            this.buttonHelp.Size = new System.Drawing.Size(75, 23);
            this.buttonHelp.TabIndex = 10;
            this.buttonHelp.Text = "Справка";
            this.buttonHelp.UseVisualStyleBackColor = true;
            this.buttonHelp.Click += new System.EventHandler(this.buttonHelp_Click);
            // 
            // groupBox4
            // 
            this.groupBox4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox4.Controls.Add(this.comboBoxDWGExportTypes);
            this.groupBox4.Controls.Add(this.label1);
            this.groupBox4.Controls.Add(this.checkBox_isDWGExport);
            this.groupBox4.Controls.Add(this.label9);
            this.groupBox4.Controls.Add(this.textBox_DWGPath);
            this.groupBox4.Controls.Add(this.buttonDWGBrowse);
            this.groupBox4.Location = new System.Drawing.Point(250, 377);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(278, 113);
            this.groupBox4.TabIndex = 1;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Экспорт в DWG";
            // 
            // comboBoxDWGExportTypes
            // 
            this.comboBoxDWGExportTypes.FormattingEnabled = true;
            this.comboBoxDWGExportTypes.IntegralHeight = false;
            this.comboBoxDWGExportTypes.Items.AddRange(new object[] {
            "<настройки...>"});
            this.comboBoxDWGExportTypes.Location = new System.Drawing.Point(120, 40);
            this.comboBoxDWGExportTypes.Name = "comboBoxDWGExportTypes";
            this.comboBoxDWGExportTypes.Size = new System.Drawing.Size(150, 21);
            this.comboBoxDWGExportTypes.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 43);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(115, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Настройки экспорта:";
            // 
            // checkBox_isDWGExport
            // 
            this.checkBox_isDWGExport.AutoSize = true;
            this.checkBox_isDWGExport.Location = new System.Drawing.Point(9, 18);
            this.checkBox_isDWGExport.Margin = new System.Windows.Forms.Padding(2);
            this.checkBox_isDWGExport.Name = "checkBox_isDWGExport";
            this.checkBox_isDWGExport.Size = new System.Drawing.Size(171, 17);
            this.checkBox_isDWGExport.TabIndex = 13;
            this.checkBox_isDWGExport.Text = "Выполнить экспорт в DWG?";
            this.checkBox_isDWGExport.UseVisualStyleBackColor = true;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Enabled = false;
            this.label9.Location = new System.Drawing.Point(6, 65);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(117, 13);
            this.label9.TabIndex = 4;
            this.label9.Text = "Путь для сохранения:";
            // 
            // textBox_DWGPath
            // 
            this.textBox_DWGPath.Location = new System.Drawing.Point(5, 83);
            this.textBox_DWGPath.Name = "textBox_DWGPath";
            this.textBox_DWGPath.Size = new System.Drawing.Size(227, 20);
            this.textBox_DWGPath.TabIndex = 5;
            // 
            // buttonDWGBrowse
            // 
            this.buttonDWGBrowse.Location = new System.Drawing.Point(238, 83);
            this.buttonDWGBrowse.Name = "buttonDWGBrowse";
            this.buttonDWGBrowse.Size = new System.Drawing.Size(27, 20);
            this.buttonDWGBrowse.TabIndex = 6;
            this.buttonDWGBrowse.Text = "...";
            this.buttonDWGBrowse.UseVisualStyleBackColor = true;
            this.buttonDWGBrowse.Click += new System.EventHandler(this.buttonDWGBrowse_Click);
            // 
            // FormPrint
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(539, 681);
            this.Controls.Add(this.groupBoxNameRules);
            this.Controls.Add(this.labelNameRulesError);
            this.Controls.Add(this.treeView1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.buttonPreview);
            this.Controls.Add(this.buttonHelp);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.pluginVersion);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "FormPrint";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "KPLN: Пакетная выдача";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBoxNameRules.ResumeLayout(false);
            this.groupBoxNameRules.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBoxNameRules;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox comboBoxPrinters;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.RadioButton radioButtonRastr;
        private System.Windows.Forms.RadioButton radioButtonVector;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.ComboBox comboBoxColors;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox comboBoxRasterQuality;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button buttonPDFBrowse;
        private System.Windows.Forms.TextBox textBox_PDFPath;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnSheetOpenNameConstructor;
        private System.Windows.Forms.TextBox textBox_SheetNameConstructor;
        private System.Windows.Forms.Label labelSheetName;
        private System.Windows.Forms.RadioButton radioButtonPDF;
        private System.Windows.Forms.RadioButton radioButtonPaper;
        private System.Windows.Forms.CheckBox checkBoxMergePdfs;
        private System.Windows.Forms.Label pluginVersion;
        private System.Windows.Forms.Label labelNameRulesError;
        private System.Windows.Forms.TreeView treeView1;
        private System.Windows.Forms.CheckBox checkBoxOrientation;
        private System.Windows.Forms.Button buttonPreview;
        private System.Windows.Forms.Button buttonHelp;
        private System.Windows.Forms.CheckBox checkBoxRefresh;
        private System.Windows.Forms.CheckBox checkBoxExcludeBorders;
        private System.Windows.Forms.Button buttonExcludesColor;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox comboBoxDWGExportTypes;
        private System.Windows.Forms.Label labelViewName;
        private System.Windows.Forms.TextBox textBox_ViewNameConstructor;
        private System.Windows.Forms.Button btnViewOpenNameConstructor;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox textBox_DWGPath;
        private System.Windows.Forms.Button buttonDWGBrowse;
        private System.Windows.Forms.CheckBox checkBox_isPDFExport;
        private System.Windows.Forms.CheckBox checkBox_isDWGExport;
    }
}
