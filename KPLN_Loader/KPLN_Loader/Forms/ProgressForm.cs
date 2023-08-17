using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Threading;

namespace KPLN_Loader.Forms
{
    public partial class ProgressForm : Form
    {
        private readonly int _totalProggres;
        private static int _currentProggres;
        private readonly System.Windows.Forms.Timer _closeTimer;
        private static object _lock = new object();
        Application _loaderApp;

        public ProgressForm(Application loaderApp, string description, int max, int current = 0)
        {
            InitializeComponent();

            _header.Text = "Приступаю к копированию";
            _description.Text = description;
            _totalProggres = max;

            _progressBar.Minimum = 0;
            _progressBar.Maximum = _totalProggres;
            _progressBar.Value = current;

            _closeTimer = new System.Windows.Forms.Timer();
            _closeTimer.Tick += CloseTimer_Tick;


            loaderApp.Progress += App_Progress;
            _loaderApp = loaderApp;

            Show();
            System.Windows.Forms.Application.DoEvents();

        }

        public void Start()
        {
            _loaderApp.DoWork();
        }

        private void App_Progress(int progress, string msg)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Application.LoadModulesProgress(App_Progress), new Object[] { progress, msg });
            }
            else
            {
                
                if (progress < _totalProggres)
                {
                    _progressBar.Value = progress;
                    _header.Text = msg;
                }
                else if (progress == _totalProggres)
                {
                    _progressBar.Value = progress;
                    _header.Text = $"Загрузка завершена! Получено {progress} из {_totalProggres} модулей";
                }
            }
            this.Update();
        }


        /// <summary>
        /// Увеличение значения на 1
        /// </summary>
        //public void Increment()
        //{
        //    _currentProggres++;

        //    Thread thread = new Thread((obj) =>
        //    {
        //        if (obj is ProgressForm progressForm)
        //        {
        //            lock (_lock)
        //            {
        //                ProgressBar progressBar = progressForm._progressBar;
        //                if (progressBar.Maximum > _currentProggres)
        //                {
        //                    // Обновляем ProgressBar через UI-поток
        //                    Invoke((MethodInvoker)delegate
        //                    {
        //                        progressBar.Value = _currentProggres;
        //                        progressForm._header.Text = $"Загружаю. Получил {_currentProggres} из {_totalProggres} модулей";
        //                    });
        //                }
        //                else if (progressBar.Maximum == _currentProggres)
        //                {
        //                    Invoke((MethodInvoker)delegate
        //                    {
        //                        progressForm._header.Text = $"Загрузка завершена! Получено {_currentProggres} из {_totalProggres} модулей";
        //                        if (ProgressForm._currentProggres != _totalProggres)
        //                            progressForm._description.Text = $"Ошибки при загрузке - см. файлы логов KPLN_Loader: c:\\temp\\KPLN_Logs\\";
        //                    });
        //                }
        //            }
        //        }
        //    });

        //    thread.Start(this);
        //}




        /// <summary>
        /// Моделирует финальное окно для пользователя
        /// </summary>
        public void UserFinalizer()
        {
            

            _header.Text = $"Загрузка завершена! Получено {_currentProggres} из {_totalProggres} модулей";
            if (_currentProggres != _totalProggres)
                _description.Text = $"Ошибки при загрузке - см. файлы логов KPLN_Loader: c:\\temp\\KPLN_Logs\\";

            Start_WindowClose();
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
