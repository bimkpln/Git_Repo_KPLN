using KPLN_OpeningHoleManager.Forms.MVVMCore_MainMenu;
using System;
using System.Windows;

namespace KPLN_OpeningHoleManager.Forms
{
    public partial class ProgressWindow : Window
    {
        private System.Windows.Forms.Timer _closeTimer;

        public ProgressWindow(ProgressInfoViewModel progressInfo)
        {
            InitializeComponent();
            DataContext = progressInfo;

            progressInfo.CompleteStatus += ProgressInfo_CompleteStatus;
        }

        private void ProgressInfo_CompleteStatus()
        {
            _closeTimer = new System.Windows.Forms.Timer
            {
                Interval = 5000
            };
            _closeTimer.Start();

            _closeTimer.Tick += CloseTimer_Tick;
        }

        private void CloseTimer_Tick(object sender, EventArgs e)
        {
            // Закрываем форму после истечения задержки.
            _closeTimer.Stop();
            Close();
        }
    }
}
