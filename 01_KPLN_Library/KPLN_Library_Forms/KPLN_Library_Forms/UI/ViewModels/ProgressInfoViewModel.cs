using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Threading;

namespace KPLN_Library_Forms.UI.ViewModels
{
    public sealed class ProgressInfoViewModel : INotifyPropertyChanged
    {
        private string _processTitle = "Выполнение операции";
        private string _mainStatus = "Выполняется...";
        private int _currentProgress;
        private int _maxProgress = 1;
        private bool _isIndeterminate;
        private volatile bool _isComplete;
        private volatile bool _isCancellationRequested;
        private volatile bool _isCancellationEnabled = true;
        private Dispatcher _dispatcher;

        public event Action CompleteStatus;
        public event PropertyChangedEventHandler PropertyChanged;


        public string ProgressCount => IsIndeterminate ? string.Empty : $"{CurrentProgress}/{MaxProgress}";

        public string ProcessTitle
        {
            get => _processTitle;
            set { _processTitle = value; OnPropertyChanged(); }
        }

        public int CurrentProgress
        {
            get => _currentProgress;
            set
            {
                _currentProgress = Math.Max(0, value);
                OnPropertyChanged();
                OnPropertyChanged(nameof(ProgressCount));
            }
        }

        public int MaxProgress
        {
            get => _maxProgress;
            set
            {
                _maxProgress = Math.Max(1, value);
                OnPropertyChanged();
                OnPropertyChanged(nameof(ProgressCount));
            }
        }

        // Имя свойства сохранено для совместимости с окном пакетной выдачи.
        public string MaintStatus
        {
            get => _mainStatus;
            set { _mainStatus = value; OnPropertyChanged(); }
        }

        public bool IsIndeterminate
        {
            get => _isIndeterminate;
            set
            {
                if (_isIndeterminate == value)
                    return;

                _isIndeterminate = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ProgressCount));
            }
        }

        public bool IsComplete
        {
            get => _isComplete;
            set
            {
                if (_isComplete == value)
                    return;

                _isComplete = value;
                if (value)
                {
                    IsIndeterminate = false;
                    MaintStatus = "Завершено";
                    CompleteStatus?.Invoke();
                }

                OnPropertyChanged();
                OnPropertyChanged(nameof(CanCancel));
            }
        }

        public bool IsCancellationRequested => _isCancellationRequested;

        public bool IsCancellationEnabled
        {
            get => _isCancellationEnabled;
            set
            {
                _isCancellationEnabled = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CanCancel));
            }
        }

        public bool CanCancel => IsCancellationEnabled && !IsComplete && !IsCancellationRequested;


        public void RequestCancellation()
        {
            if (!CanCancel)
                return;

            _isCancellationRequested = true;
            MaintStatus = "Отменено пользователем";
            OnPropertyChanged(nameof(IsCancellationRequested));
            OnPropertyChanged(nameof(CanCancel));
            CompleteStatus?.Invoke();
        }

        internal void AttachDispatcher(Dispatcher dispatcher)
        {
            _dispatcher = dispatcher;
        }

        private void OnPropertyChanged([CallerMemberName] string name = null)
        {
            Dispatcher dispatcher = _dispatcher;
            if (dispatcher == null || dispatcher.CheckAccess())
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
                return;
            }

            if (!dispatcher.HasShutdownStarted && !dispatcher.HasShutdownFinished)
                dispatcher.BeginInvoke(DispatcherPriority.DataBind,
                    new Action(() => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name))));
        }
    }
}
