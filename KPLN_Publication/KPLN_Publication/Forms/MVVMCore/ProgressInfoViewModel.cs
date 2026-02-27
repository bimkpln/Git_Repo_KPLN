using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace KPLN_Publication.Forms.MVVMCore
{
    public sealed class ProgressInfoViewModel : INotifyPropertyChanged
    {
        public delegate void RiseComplete();
        public event RiseComplete CompleteStatus;

        public event PropertyChangedEventHandler PropertyChanged;

        public string ProgressCount => $"{CurrentProgress}/{MaxProgress}";

        private string _processTitle = "<Имя процесса установится по ходу анализа>";
        private int _currentProgress;
        private int _maxProgress = 1;

        private string _mainStatus = "Экспортирую.....";
        public bool _isComplete = false;
        private bool _isCancellationRequested;

        public ProgressInfoViewModel()
        {
        }

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
                _currentProgress = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ProgressCount));
            }
        }

        public int MaxProgress
        {
            get => _maxProgress;
            set
            {
                _maxProgress = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ProgressCount));
            }
        }

        public string MaintStatus
        {
            get => _mainStatus;
            set { _mainStatus = value; OnPropertyChanged(); }
        }

        public bool IsComplete
        {
            get => _isComplete;
            set
            {
                _isComplete = value;
                if (_isComplete)
                {
                    MaintStatus = "Завершено";
                    CompleteStatus?.Invoke();
                }
                OnPropertyChanged();
            }
        }

        public bool IsCancellationRequested
        {
            get => _isCancellationRequested;
            private set
            {
                _isCancellationRequested = value;
                OnPropertyChanged();
            }
        }

        public void RequestCancellation()
        {
            if (IsCancellationRequested || IsComplete)
                return;

            IsCancellationRequested = true;
            MaintStatus = "Отменено пользователем";
            CompleteStatus?.Invoke();
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

        private void OnPropertyChanged([CallerMemberName] string name = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
