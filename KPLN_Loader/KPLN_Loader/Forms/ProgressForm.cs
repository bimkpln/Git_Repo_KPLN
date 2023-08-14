using System;
using System.Windows.Forms;

namespace KPLN_Loader.Forms
{
    public partial class ProgressForm : Form
    {
        private readonly int _totalProggres;
        private int _currentProggres;
        private readonly Timer _closeTimer;

        public ProgressForm(string description, int max, int current = 0)
        {
            InitializeComponent();

            _description.Text = description;
            _totalProggres = max;

            _progressBar.Minimum = 0;
            _progressBar.Maximum = _totalProggres;
            _progressBar.Value = current;

            _closeTimer = new Timer();
            _closeTimer.Tick += CloseTimer_Tick;

            Show();
            System.Windows.Forms.Application.DoEvents();
        }

        /// <summary>
        /// Увеличение значения на 1
        /// </summary>
        public void Increment()
        {
            if (_progressBar.Maximum >= _progressBar.Value)
            {
                _currentProggres++;
                _progressBar.Value = _currentProggres;
            }
            _header.Text = $"Загружаю. Получил {_currentProggres} из {_totalProggres} модулей";

            System.Windows.Forms.Application.DoEvents();
        }

        /// <summary>
        /// Моделирует финальное окно для пользователя
        /// </summary>
        public void UserFinalizer()
        {
            _header.Text = $"Загрузка завершена! Получено {_currentProggres} из {_totalProggres} модулей";

            if (_currentProggres != _totalProggres)
                _description.Text = $"Ошибки при загрузке - см. файлы логов KPLN_Loader: c:\\temp\\KPLN_Logs\\";

            Start_WindowClose();
            System.Windows.Forms.Application.DoEvents();
        }


        /// <summary>
        /// Стартер закрытия окна с отложенным сроком
        /// </summary>
        private void Start_WindowClose()
        {
            // Запускаем таймер при загрузке формы, чтобы окно провисело гарантировано 10 сек.
            _closeTimer.Interval = 10000;
            _closeTimer.Start();
        }

        private void CloseTimer_Tick(object sender, EventArgs e)
        {
            // Закрываем форму после истечения задержки.
            _closeTimer.Stop();
            Close();
        }
    }
}
