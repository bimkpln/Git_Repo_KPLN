using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KPLN_Parameters_Ribbon.Forms
{
    public partial class Progress_Single : Form
    {
        public string _format;
        public Progress_Single(string header, string format, int max, int current = 0)
        {
            _format = format;
            InitializeComponent();
            Text = header;
            Header_lbl.Text = (null == format) ? header : string.Format(format, 0);
            Add_lbl.Text = "инициализация...";
            Titile_lbl.Text = header;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = max;
            progressBar1.Value = current;
            Show();
            System.Windows.Forms.Application.DoEvents();
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
            if (value != null)
            {
                Add_lbl.Text = value;
            }
            progressBar1.Value = progressvalue;
            if (null != _format)
            {
                Header_lbl.Text = string.Format(_format, progressBar1.Value);
            }
            System.Windows.Forms.Application.DoEvents();
        }
    }
}
