using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Threading;

namespace KPLN_OpeningHoleManager.Forms.MVVMCore_MainMenu
{
    public sealed class ProgressInfoViewModel : INotifyPropertyChanged
    {
        public delegate void RiseComplete();
        public event RiseComplete CompleteStatus;
        
        public event PropertyChangedEventHandler PropertyChanged;

        private string _processTitle = "<Имя процесса установится по ходу анализа>";
        private int _currentProgress;
        private int _maxProgress = 1;

        private string _mainStatus = "Анализирую.....";
        public bool _isComplete = false;

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
            set { _currentProgress = value; OnPropertyChanged(); }
        }

        public int MaxProgress
        {
            get => _maxProgress;
            set { _maxProgress = value; OnPropertyChanged(); }
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
                    MaintStatus = "Завершено!";
                    CompleteStatus?.Invoke();
                }

                OnPropertyChanged(); 
            }
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
