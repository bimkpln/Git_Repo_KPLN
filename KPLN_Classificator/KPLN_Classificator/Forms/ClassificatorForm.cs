using KPLN_Classificator.Data;
using KPLN_Classificator.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using static KPLN_Classificator.ApplicationConfig;
using BuiltInCategory = Autodesk.Revit.DB.BuiltInCategory;

namespace KPLN_Classificator.Forms
{
    public partial class ClassificatorForm : Form
    {
        public StorageUtils utils;
        public InfosStorage storage;
        public bool debugMode;
        public bool colourMode;
        public List<BuiltInCategory> checkedCats;
        public MainWindow form;
        public List<BuiltInCategory> allCats;
        public List<MyParameter> mparams;
        public long hashSumOfClassificator;

        public LastRunInfo lastRunInfo;

        public ClassificatorForm(StorageUtils utils, List<MyParameter> mparams)
        {
            if (IsDocumentAvailable)
                this.lastRunInfo = LastRunInfo.getInstance();
            
            initForm(utils, mparams);
        }

        private void initForm(StorageUtils utils, List<MyParameter> mparams)
        {
            InitializeComponent();
            this.utils = utils;
            debugMode = checkBoxDebug.Checked;
            colourMode = checkBoxColour.Checked;
            textBoxFileInfo.Text = "Конфигурационный файл не выбран.";
            textBoxFileInfo.Text += Environment.NewLine;
            btnRun.Enabled = false;
            checkedCats = new List<BuiltInCategory>();
            buttonOpenConfiguration.Enabled = false;
            buttonSaveFile.Enabled = false;
            this.mparams = mparams;
            
            if (IsDocumentAvailable && lastRunInfo != null)
                getInfoAboutFileClassification();
            
            if (lastRunInfo != null && System.IO.File.Exists(lastRunInfo.getFileName()))
                buttonChooseLastFile.Enabled = true;
            else
                buttonChooseLastFile.Enabled = false;
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            if (ApplicationConfig.IsDocumentAvailable)
            {
                foreach (var checkedItem in checkedListBox1.CheckedItems)
                {
                    BuiltInCategory cat = (BuiltInCategory)checkedItem;
                    checkedCats.Add(cat);
                }
                CurrentCmdEnv.toEnqueue(new CommandStartClassificator(this));
            }
            else
            {
                MessageBox.Show("Ни один документ не открыт. Запуск классификатора невозможен.", "Ошибка!", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void buttonChooseFile_Click(object sender, EventArgs e)
        {
            storage = utils.getInfoStorage();
            if (storage != null)
            {
                if (hashSumOfClassificator != 0)
                {
                    hashSumOfClassificator = 0L;
                    form = null;
                }
                checkFileIsCorrect();
                setRadioButton();
            }
        }

        private void buttonChooseLastFile_Click(object sender, EventArgs e)
        {
            storage = utils.getStorageFromFilePath(lastRunInfo.getFileName());
            if (storage != null)
            {
                if (hashSumOfClassificator != 0)
                {
                    hashSumOfClassificator = 0L;
                    form = null;
                }
                checkFileIsCorrect();
                setRadioButton();
            }
        }

        private void checkBoxDebug_CheckedChanged(object sender, EventArgs e)
        {
            debugMode = checkBoxDebug.Checked;
        }

        private void checkBoxColour_CheckedChanged(object sender, EventArgs e)
        {
            colourMode = checkBoxColour.Checked;
        }

        private bool checkFileIsCorrect()
        {
            if ((storage.instanceOrType == 1) || (storage.instanceOrType == 2))
            {
                textBoxFileInfo.Text = "";
                textBoxFileInfo.Text += "Выбран конфигурационный файл:" + Environment.NewLine;
                textBoxFileInfo.Text += Environment.NewLine;
                textBoxFileInfo.Text += hashSumOfClassificator == getHashSumOfClassificator()
                    ? (utils.xmlFilePath != null ? utils.xmlFilePath : "") + Environment.NewLine :
                    utils.xmlFilePath != null ? utils.xmlFilePath + Environment.NewLine + "Внимание! Конфигурационный файл изменён и не сохранён!" + Environment.NewLine
                    : "Внимание! Конфигурационный файл изменён и не сохранён!" + Environment.NewLine;
                textBoxFileInfo.Text += Environment.NewLine;
                string info = storage.instanceOrType == 1 ? "ЭКЗЕМПЛЯРУ" : storage.instanceOrType == 2 ? "ТИПУ" : "НЕКОРРЕКТНАЯ НАСТРОЙКА КОНФИГУРАЦИОННОГО ФАЙЛА!";
                textBoxFileInfo.Text += string.Format("Файл содержит {0} правил(о/а) для заполнения классификатора по {1}.", storage.classificator.Count.ToString(), info);
                textBoxFileInfo.Text += Environment.NewLine;
                if (ApplicationConfig.IsDocumentAvailable && lastRunInfo != null)
                {
                    getInfoAboutFileClassification();
                }
                btnRun.Enabled = true;

                checkedListBox1.Items.Clear();
                foreach (BuiltInCategory bic in storage.classificator.Select(x => x.BuiltInName).ToHashSet().ToList())
                {
                    checkedListBox1.Items.Add(bic, CheckState.Checked);
                }
                buttonOpenConfiguration.Enabled = true;

                if (System.IO.File.Exists(lastRunInfo.getFileName()))
                {
                    buttonChooseLastFile.Enabled = true;
                }
                else
                {
                    buttonChooseLastFile.Enabled = false;
                }

                return true;
            }
            else
            {
                checkedListBox1.Items.Clear();
                textBoxFileInfo.Text = "";
                textBoxFileInfo.Text += "НЕКОРРЕКТНАЯ НАСТРОЙКА КОНФИГУРАЦИОННОГО ФАЙЛА!";
                btnRun.Enabled = false;
                buttonSaveFile.Enabled = false;
                return false;
            }
        }

        private void getInfoAboutFileClassification()
        {
            string[] infoArray = lastRunInfo.ToString().Split('=');
            textBoxFileInfo.Text += Environment.NewLine;
            foreach (var item in infoArray)
            {
                textBoxFileInfo.Text += item + Environment.NewLine;
            }
        }

        private void setRadioButton()
        {
            if (storage != null && storage.instanceOrType == 1)
                buttonSaveFile.Enabled = true;
            else if (storage != null && storage.instanceOrType == 2)
                buttonSaveFile.Enabled = true;
        }

        private void buttonOpenConfiguration_Click(object sender, EventArgs e)
        {
            if (allCats == null)
            {
                getBuiltInCategorys();
            }
            this.Hide();
            form = new MainWindow(this, storage, allCats);
            if (form != null && hashSumOfClassificator == 0)
            {
                hashSumOfClassificator = getHashSumOfClassificator();
            }
            form.Show();
        }

        private void ClassificatorForm_VisibleChanged(object sender, EventArgs e)
        {
            if (storage != null && storage.instanceOrType != 0)
            {
                checkFileIsCorrect();
                setRadioButton();
                buttonOpenConfiguration.Enabled = true;
                buttonSaveFile.Enabled = true;
            }
            else
            {
                buttonOpenConfiguration.Enabled = false;
                buttonSaveFile.Enabled = false;
            }
        }

        private void buttonCreateNewConfiguration_Click(object sender, EventArgs e)
        {
            if (storage != null && storage.instanceOrType != 0 && MessageBox.Show("При создании нового конфигурационного файла, старый будет удалён!", "Внимание!", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.Cancel)
            {
                return;
            }

            if (allCats == null)
            {
                getBuiltInCategorys();
            }
            utils.xmlFilePath = null;
            this.Hide();
            storage = new InfosStorage();
            form = new MainWindow(this, storage, allCats);
            form.Show();
        }

        private void getBuiltInCategorys()
        {
            allCats = new List<BuiltInCategory>();
            foreach (var item in Enum.GetValues(typeof(BuiltInCategory)))
            {
                allCats.Add((BuiltInCategory)item);
            }
            allCats.Sort((x, y) => x.ToString().CompareTo(y.ToString()));
        }

        private void buttonSaveFile_Click(object sender, EventArgs e)
        {
            utils.saveInfoStorage(storage);
            hashSumOfClassificator = getHashSumOfClassificator();
            checkFileIsCorrect();
        }

        private long getHashSumOfClassificator()
        {
            long sum = 0L;
            if (form != null && form.ruleItems != null)
            {
                foreach (var ruleItem in form.ruleItems)
                {
                    sum += ruleItem.GetHashCode();
                    foreach (var valueItem in ruleItem.valuesOfParams)
                    {
                        sum += valueItem.GetHashCode();
                    }
                }
                foreach (var paramName in form.settings.paramNameItems)
                {
                    sum += paramName.GetHashCode();
                }
            }
            return sum;
        }

        private void ClassificatorForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (hashSumOfClassificator != getHashSumOfClassificator())
            {
                if (MessageBox.Show("Закрыть окно? Изменения будут потеряны!", "Внимание!", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) != DialogResult.OK)
                {
                    e.Cancel = true;
                }
            }
        }
    }
}
