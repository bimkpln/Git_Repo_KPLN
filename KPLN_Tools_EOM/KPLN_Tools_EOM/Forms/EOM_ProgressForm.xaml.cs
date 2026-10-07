using KPLN_Library_Forms.Services;
using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Threading;

namespace KPLN_Tools_EOM.Forms
{
    public partial class EOM_ProgressForm : Window
    {
        private bool _completed;

        public EOM_ProgressForm()
        {
            InitializeComponent();
            WindowHandleSearch.MainWindowHandle.SetAsOwner(this);
        }

        public void UpdateProgress(int value, string statusText)
        {
            Progress.Value = Math.Max(0, Math.Min(value, 100));
            StatusText.Text = statusText;
            Dispatcher.Invoke(DispatcherPriority.Render, new Action(() => { }));
        }

        public void Complete()
        {
            _completed = true;
            Close();
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            e.Cancel = !_completed;
            base.OnClosing(e);
        }
    }
}
