using System.Windows.Threading;

namespace KPLN_OpeningHoleManager.Forms.MVVMCore_MainMenu
{
    /// <summary>
    /// Адаптер существующих операций менеджера к библиотечному окну прогресса.
    /// </summary>
    public sealed class ProgressInfoViewModel
    {
        internal KPLN_Library_Forms.UI.ViewModels.ProgressInfoViewModel LibraryProgress { get; }
            = new KPLN_Library_Forms.UI.ViewModels.ProgressInfoViewModel
            {
                IsCancellationEnabled = false
            };

        public string ProcessTitle
        {
            get => LibraryProgress.ProcessTitle;
            set => LibraryProgress.ProcessTitle = value;
        }

        public int CurrentProgress
        {
            get => LibraryProgress.CurrentProgress;
            set => LibraryProgress.CurrentProgress = value;
        }

        public int MaxProgress
        {
            get => LibraryProgress.MaxProgress;
            set => LibraryProgress.MaxProgress = value;
        }

        public string MaintStatus
        {
            get => LibraryProgress.MaintStatus;
            set => LibraryProgress.MaintStatus = value;
        }

        public bool IsComplete
        {
            get => LibraryProgress.IsComplete;
            set => LibraryProgress.IsComplete = value;
        }

        public void DoEvents()
        {
            DispatcherFrame frame = new DispatcherFrame();

            Dispatcher.CurrentDispatcher.BeginInvoke(
                DispatcherPriority.Background,
                new DispatcherOperationCallback(f =>
                {
                    ((DispatcherFrame)f).Continue = false;
                    return null;
                }),
                frame);

            Dispatcher.PushFrame(frame);
        }
    }
}
