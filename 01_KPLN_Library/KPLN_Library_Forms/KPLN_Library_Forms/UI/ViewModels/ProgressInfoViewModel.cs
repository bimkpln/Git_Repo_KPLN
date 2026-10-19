using System;
using System.ComponentModel;
using System.Diagnostics;
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
        private readonly Stopwatch _reportClock = Stopwatch.StartNew();
        private string _lastReportedStage;

        /// <summary>Минимальный интервал обновлений внутри этапа. Вызывать Report из потока операции.</summary>
        public int UpdateIntervalMilliseconds { get; set; } = 75;

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

        /// <summary>Обновляет счётчик с ограничением частоты; отмена проверяется при каждом вызове.</summary>
        public void Report(int completed, int total, string stage)
        {
            ThrowIfCancellationRequested();
            if (!IsIndeterminate && stage == _lastReportedStage && total == MaxProgress
                && completed > 0 && completed < total
                && _reportClock.ElapsedMilliseconds < UpdateIntervalMilliseconds)
                return;

            _lastReportedStage = stage;
            _reportClock.Restart();
            MaintStatus = stage;
            MaxProgress = total;
            CurrentProgress = completed;
            IsIndeterminate = false;
        }

        /// <summary>Этап, для которого пока неизвестно число операций.</summary>
        public void SetStage(string stage)
        {
            ThrowIfCancellationRequested();
            _lastReportedStage = null;
            MaintStatus = stage;
            IsIndeterminate = true;
        }

        public void ThrowIfCancellationRequested()
        {
            if (IsCancellationRequested)
                throw new OperationCanceledException();
        }

        /// <summary>Отключает отмену перед окончательным сохранением; принятую отмену не игнорирует.</summary>
        public void BeginCommit()
        {
            IsCancellationEnabled = false;
            SetStage("Сохраняю изменения");
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
