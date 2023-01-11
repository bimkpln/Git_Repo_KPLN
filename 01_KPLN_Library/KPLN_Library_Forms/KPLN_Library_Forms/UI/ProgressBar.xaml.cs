using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Input;

namespace KPLN_Library_Forms.UI
{
    public partial class ProgressBar : Window, INotifyPropertyChanged
    {
        private bool _isHeaderVisible;

        private int _currentProgress;

        private int _itemMaxCount;

        private int _progressStep;

        private BackgroundWorker _backgroundWorker;

        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Прогресс бар по процессу
        /// </summary>
        /// <param name="itemMaxCount">Количество элементов, которые соответсвуют 100%</param>
        /// <param name="title">Имя окна, в формате "KPLN: ..."</param>
        /// <param name="mainDescription">Описание основной задачи</param>
        /// <param name="isHeaderVisible">Наличие хэдера (сворачиваемой подсказки)</param>
        /// <param name="headerTitle">Имя хэдера (если есть)</param>
        /// <param name="headerDescription">Содержание хэдера (если есть)</param>
        public ProgressBar(int itemMaxCount, string title, string mainDescription, bool isHeaderVisible, string headerTitle = null, string headerDescription = null)
        {
            InitializeComponent();

            _itemMaxCount = itemMaxCount;
            _progressStep = 100 / _itemMaxCount;
            _currentProgress = 0;

            Title = $"KPLN: {title}";
            tblMainDescription.Text = mainDescription;
            IsHeaderVisible = isHeaderVisible;
            expHeader.Header = headerTitle;
            tblHeaderDescription.Text = headerDescription;

            InitializeBackgroundWorker();
        }

        public bool IsHeaderVisible
        {
            get { return _isHeaderVisible; }
            private set { _isHeaderVisible = value; }
        }

        public int CurrentProgress
        {
            get { return _currentProgress; }
            set 
            { 
                _currentProgress = value * _progressStep;
                //PropertyChanged(() => CurrentProgress);
            }
        }

        private void InitializeBackgroundWorker()
        {
            _backgroundWorker = new BackgroundWorker();

            _backgroundWorker.DoWork +=
                new DoWorkEventHandler(Worker_DoWork);

            _backgroundWorker.ProgressChanged +=
                new ProgressChangedEventHandler(Worker_ProgressChenged);

            _backgroundWorker.RunWorkerCompleted +=
                new RunWorkerCompletedEventHandler(Worker_RunWorkerComplite);
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            if (CurrentProgress < 100 || CurrentProgress == 100)
                worker.ReportProgress(CurrentProgress);
            else
                worker.ReportProgress(100);
        }

        private void Worker_ProgressChenged(object sender, ProgressChangedEventArgs e)
        {
            pbStatus.Value = e.ProgressPercentage;
        }

        private void Worker_RunWorkerComplite(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
                Close();
        }
    }
}
