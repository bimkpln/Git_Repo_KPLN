using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using KPLN_Views_Ribbon.Views.FilterUtils;
using KPLN_Views_Ribbon.Views.Colorize;


namespace KPLN_Views_Ribbon.Views.Colorize
{

    public partial class FormSelectParameterForFilters : Form
    {
        public List<MyParameter> parameters;
        public MyParameter selectedParameter;
        public CriteriaType criteriaType;
        public ColorizeMode colorizeMode;
        public int startSymbols;
        public bool colorLines;
        public bool colorFill;


        public FormSelectParameterForFilters()
        {
            InitializeComponent();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void FormSelectParameterForFilters_Load(object sender, EventArgs e)
        {
            comboBoxParameters.DataSource = parameters;
        }

        private void buttonNext_Click(object sender, EventArgs e)
        {
            if(radioButtonCheckHostMark.Checked)
            {
                this.colorizeMode = ColorizeMode.CheckHostmark;
            }

            if (radioButtonUserParameter.Checked)
            {
                this.colorizeMode = ColorizeMode.ByParameter;
                selectedParameter = comboBoxParameters.SelectedItem as MyParameter;
                if (selectedParameter == null) throw new Exception("Выбран не MyParameter");


                if (radioButtonEquals.Checked) criteriaType = CriteriaType.Equals;
                if (radioButtonStartsWith.Checked) criteriaType = CriteriaType.StartsWith;
                startSymbols = (int)numericStartSymbols.Value;
            }

            this.colorLines = checkBoxColorLines.Checked;
            this.colorFill = checkBoxColorFill.Checked;

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void radioButtonEquals_CheckedChanged(object sender, EventArgs e)
        {
            numericStartSymbols.Enabled = false;
            labelSymbols.Enabled = false;
        }

        private void radioButtonStartsWith_CheckedChanged(object sender, EventArgs e)
        {
            numericStartSymbols.Enabled = true;
            labelSymbols.Enabled = true;

        }

        private void radioButtonUserParameter_CheckedChanged(object sender, EventArgs e)
        {
            groupBox1.Enabled = true;
        }

        private void radioButtonCheckHostMark_CheckedChanged(object sender, EventArgs e)
        {
            groupBox1.Enabled = false;
        }
    }
}
