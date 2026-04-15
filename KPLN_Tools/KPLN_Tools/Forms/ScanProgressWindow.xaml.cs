using System.Windows;

namespace KPLN_Tools.Forms
{
    public partial class ScanProgressWindow : Window
    {
        public ScanProgressWindow()
        {
            InitializeComponent();
        }

        public void UpdateProgress(int current, int total, string status)
        {
            StatusText.Text = status;
            PercentText.Text = current + " / " + total;

            if (total <= 0)
            {
                MainProgressBar.Value = 0;
                return;
            }

            double percent = (double)current / total * 100.0;
            MainProgressBar.Value = percent;
        }
    }
}