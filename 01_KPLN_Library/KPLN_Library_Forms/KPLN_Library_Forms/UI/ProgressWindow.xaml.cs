using KPLN_Library_Forms.UI.ViewModels;
using System;
using System.ComponentModel;
using System.Threading;
using System.Windows;
using System.Windows.Threading;
using System.Windows.Interop;

namespace KPLN_Library_Forms.UI
{
    public partial class ProgressWindow : Window
    {
        private readonly ProgressInfoViewModel _progressInfo;
        private DispatcherTimer _closeTimer;
        private bool _ownsDispatcher;
        private bool _programmaticClose;
        private bool _isClosed;

        public ProgressWindow(ProgressInfoViewModel progressInfo)
        {
            _progressInfo = progressInfo ?? throw new ArgumentNullException(nameof(progressInfo));
            InitializeComponent();
            DataContext = progressInfo;
            progressInfo.AttachDispatcher(Dispatcher);
            progressInfo.CompleteStatus += ProgressInfo_CompleteStatus;
            Closing += OnClosing;
            Closed += OnClosed;
        }


        /// <summary>
        /// Только окно работает в отдельном STA-потоке. Revit API остаётся в вызывающем потоке.
        /// </summary>
        public static ProgressWindow ShowInBackground(ProgressInfoViewModel progressInfo,
            string title = "KPLN: Выполнение")
        {
            // Сигнатура и прежнее поведение сохранены для уже собранных плагинов.
            return ShowInBackground(progressInfo, title, true, IntPtr.Zero);
        }

        public static ProgressWindow ShowInBackground(ProgressInfoViewModel progressInfo,
            string title, bool topmost, IntPtr ownerHandle)
        {
            ProgressWindow window = null;
            Exception startupError = null;
            using (var ready = new ManualResetEventSlim(false))
            {
                var thread = new Thread(() =>
                {
                    try
                    {
                        window = new ProgressWindow(progressInfo) { Title = title, _ownsDispatcher = true };
                        window.Topmost = topmost;
                        if (ownerHandle != IntPtr.Zero)
                            new WindowInteropHelper(window).Owner = ownerHandle;
                        window.Show();
                    }
                    catch (Exception ex)
                    {
                        startupError = ex;
                        window?.CloseSafely();
                    }
                    finally
                    {
                        ready.Set();
                    }

                    if (startupError == null)
                        System.Windows.Threading.Dispatcher.Run();
                });
                thread.SetApartmentState(ApartmentState.STA);
                thread.IsBackground = true;
                thread.Start();
                ready.Wait();
            }

            if (startupError != null)
                throw new InvalidOperationException("Не удалось открыть окно прогресса.", startupError);

            return window;
        }

        public void CloseSafely()
        {
            if (Dispatcher.HasShutdownStarted || Dispatcher.HasShutdownFinished)
                return;

            Action close = () =>
            {
                if (_isClosed)
                    return;

                _programmaticClose = true;
                Close();
            };

            if (Dispatcher.CheckAccess())
                close();
            else
            {
                try
                {
                    Dispatcher.Invoke(close);
                }
                catch (OperationCanceledException) when (Dispatcher.HasShutdownStarted || Dispatcher.HasShutdownFinished)
                {
                    // Окно уже закрыто пользователем, его поток завершает работу.
                }
                catch (InvalidOperationException) when (Dispatcher.HasShutdownStarted || Dispatcher.HasShutdownFinished)
                {
                    // Завершение Dispatcher могло начаться между проверкой и Invoke.
                }
            }
        }


        private void ProgressInfo_CompleteStatus()
        {
            if (!Dispatcher.CheckAccess())
            {
                if (!Dispatcher.HasShutdownStarted && !Dispatcher.HasShutdownFinished)
                    Dispatcher.BeginInvoke(new Action(ProgressInfo_CompleteStatus));
                return;
            }

            if (_isClosed || _closeTimer != null)
                return;

            _closeTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(5) };
            _closeTimer.Tick += CloseTimer_Tick;
            _closeTimer.Start();
        }

        private void CloseTimer_Tick(object sender, EventArgs e)
        {
            CloseSafely();
        }

        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            _progressInfo.RequestCancellation();
        }

        private void OnClosing(object sender, CancelEventArgs e)
        {
            if (!_programmaticClose && !_progressInfo.IsComplete && _progressInfo.WaitForCancellationCompletion)
            {
                e.Cancel = true;
                _progressInfo.RequestCancellation();
                return;
            }

            if (_programmaticClose || _progressInfo.IsComplete || _progressInfo.IsCancellationRequested)
                return;

            if (!_progressInfo.IsCancellationEnabled)
                e.Cancel = true;
            else
                _progressInfo.RequestCancellation();
        }

        private void OnClosed(object sender, EventArgs e)
        {
            _isClosed = true;
            _progressInfo.CompleteStatus -= ProgressInfo_CompleteStatus;
            if (_closeTimer != null)
            {
                _closeTimer.Stop();
                _closeTimer.Tick -= CloseTimer_Tick;
            }

            if (_ownsDispatcher)
                Dispatcher.BeginInvokeShutdown(DispatcherPriority.Background);
        }
    }
}
