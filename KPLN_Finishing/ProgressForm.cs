using System.Windows.Forms;

namespace KPLN_Finishing
{
    public partial class ProgressForm : Form
    {
        public string _format;
        public ProgressForm(string caption, string format, int max, int lvlmax)
        {
            _format = format;
            InitializeComponent();
            Text = caption;
            label1.Text = (null == format) ? caption : string.Format(format, 0);
            label2.Text = "инициализация...";
            progressBar1.Minimum = 0;
            progressBar2.Minimum = 0;
            progressBar1.Maximum = max;
            progressBar2.Maximum = lvlmax;
            progressBar1.Value = 0;
            progressBar2.Value = 0;
            Show();
            System.Windows.Forms.Application.DoEvents();
        }
        public void Increment()
        {
            ++progressBar1.Value;
            if (null != _format)
            {
                label1.Text = string.Format(_format, progressBar1.Value);
            }
            System.Windows.Forms.Application.DoEvents();
        }
        public void ResetMax(int max = 100)
        {
            progressBar1.Value = 0;
            progressBar1.Maximum = max;
            label1.Text = string.Format(_format, progressBar1.Value);
            if (null != _format)
            {
                label1.Text = string.Format(_format, progressBar1.Value);
            }
            System.Windows.Forms.Application.DoEvents();
        }
        public void SetInfoStrip(string value)
        {
            label2.Text = value;
            System.Windows.Forms.Application.DoEvents();
        }
        public void SetInfoLevel(string value)
        {
            ++progressBar2.Value;
            label3.Text = value;
            System.Windows.Forms.Application.DoEvents();
        }
    }
}
