using System.Windows;

namespace KPLN_Tools_SS.Forms
{
    public partial class ScanProgressWindow : Window
    {
        public ScanProgressWindow()
        {
            InitializeComponent();
            KPLN_Library_Forms.Services.WindowHandleSearch.MainWindowHandle.SetAsOwner(this);
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