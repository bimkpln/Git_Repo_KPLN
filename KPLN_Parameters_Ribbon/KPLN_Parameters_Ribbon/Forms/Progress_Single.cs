using System.Windows.Forms;

namespace KPLN_Parameters_Ribbon.Forms
{
    public partial class Progress_Single : Form
    {
        private string _format;
        
        public Progress_Single(string header, string format, bool isBtnOkVisible)
        {
            _format = format;
            InitializeComponent();
            
            Text = header;
            Header_lbl.Text = (null == format) ? header : string.Format(format, 0);
            Titile_lbl.Text = header;
            this.Btn_Ok.Visible = isBtnOkVisible;
        }

        public void SetProggresValues(int max, int current)
        {
            progressBar1.Minimum = 0;
            progressBar1.Maximum = max;
            progressBar1.Value = current;
        }

        public void ShowProgress()
        {
            this.Show();
            Application.DoEvents();
        }

        /// <summary>
        /// Увеличение значения на 1
        /// </summary>
        /// <param name="value"></param>
        public void Increment(string value = null)
        {
            if (value != null)
            {
                Add_lbl.Text = value;
            }
            if (progressBar1.Maximum > progressBar1.Value)
            {
                ++progressBar1.Value;
                if (null != _format)
                {
                    Header_lbl.Text = string.Format(_format, progressBar1.Value);
                }
            }
            System.Windows.Forms.Application.DoEvents();
        }

        /// <summary>
        /// Увеличение значения на 1
        /// </summary>
        /// <param name="value"></param>
        public void AddProgress(int value)
        {
            progressBar1.Value += value;
            Header_lbl.Text = string.Format(_format, progressBar1.Value);
            System.Windows.Forms.Application.DoEvents();
        }

        /// <summary>
        /// Уменьшение значения на 1, или на определенное значение
        /// </summary>
        public void Decrement(int value = 1)
        {
            if (progressBar1.Value > value)
                progressBar1.Value -= value;
            else
                progressBar1.Value = 0;

            if (null != _format)
            {
                Header_lbl.Text = string.Format(_format, progressBar1.Value);
            }

            System.Windows.Forms.Application.DoEvents();
        }


        public void Update(int progressvalue, string value = null)
        {
            if (!string.IsNullOrEmpty(value) && !Add_lbl.Text.Equals(value))
                Add_lbl.Text = value;
            
            progressBar1.Value = progressvalue;

            if (!string.IsNullOrEmpty(_format))
                Header_lbl.Text = string.Format(_format, progressvalue);

            System.Windows.Forms.Application.DoEvents();
        }

        /// <summary>
        /// Метод для отображения возможности закрытия окна, чтобы раньше времени не закрыть процесс
        /// </summary>
        public void SetBtn_Ok_Enabled()
        {
            this.Btn_Ok.Enabled = true;
        }

        private void Btn_Ok_Click(object sender, System.EventArgs e)
        {
            this.Close();
        }
    }
}
