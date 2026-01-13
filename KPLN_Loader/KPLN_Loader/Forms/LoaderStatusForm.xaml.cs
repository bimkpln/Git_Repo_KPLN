using KPLN_Loader.Core.Entities;
using KPLN_Loader.Forms.Common;
using KPLN_Loader.Services;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace KPLN_Loader.Forms
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class LoaderStatusForm : Window
    {

        internal delegate void RiseLikeEvant(MainDB_LoaderDescriptions_RateType rateType, LoaderDescription loaderDescription);
        /// <summary>
        /// Событие, которое посылает сигналы из формы, в случае активности пользователя
        /// </summary>
        internal event RiseLikeEvant LikeStatus;

        private readonly string _statusError = "❌";
        private readonly string _statusDone = "✔️";
        private readonly IEnumerable<LoaderStatusEntity> _loaderStatusEntitys;
        private readonly List<LoaderEvantEntity> _loadModules;
        private readonly System.Windows.Forms.Timer _closeTimer;
        private LoaderDescription _loaderDescription;

        /// <summary>
        /// Окно с демонстрацией прогресса
        /// </summary>
        internal LoaderStatusForm(Application kplnApp)
        {
            kplnApp.Progress += Application_Progress;
            kplnApp.LoadStatus += Application_ModuleStatus;

            _closeTimer = new System.Windows.Forms.Timer();
            _closeTimer.Tick += CloseTimer_Tick;

            _loaderStatusEntitys = new List<LoaderStatusEntity>
            {
                new LoaderStatusEntity("Подготовка окружения", _statusError, MainStatus.Envirnment),
                new LoaderStatusEntity("Подключение к базам данных", _statusError, MainStatus.DbConnection),
                new LoaderStatusEntity("Активация модулей", _statusError, MainStatus.ModulesActivation)
            };
            _loadModules = new List<LoaderEvantEntity>();

            InitializeComponent();

            versionTxt.Text = "v." + Assembly.GetExecutingAssembly().GetName().Version.ToString();
            stepListBox.ItemsSource = _loaderStatusEntitys;
            DataContext = this;
        }

        /// <summary>
        /// Текстовое описание
        /// </summary>
        internal void SetInstruction(LoaderDescription loaderDescription)
        {
            Dispatcher.Invoke(() =>
            {
                _loaderDescription = loaderDescription;
                tblInstruction.Text = _loaderDescription.Description;

            string loaderDescriptionURL = _loaderDescription.InstructionURL;
            if (loaderDescriptionURL != null)
            {
                tblInstruction.TextDecorations = TextDecorations.Underline;
                tblInstruction.Foreground = new SolidColorBrush(Colors.Blue);
            }
        });
        }

        /// <summary>
        /// Добавляет маркер при дебаге модулей
        /// </summary>
        /// <param name="isDebugMode">Указатель</param>
        internal void SetDebugStatus(bool isDebugMode)
        {
            Dispatcher.Invoke(() =>
            {
                if (isDebugMode)
                    DebugModeTxt.Visibility = Visibility.Visible;
            });
        }

        /// <summary>
        /// Событие наведения мыши в окно ScrollViewer. Связано с потеряй фокуса на колесо мыши
        /// </summary>
        private void ScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            // Выполните здесь нужное вам поведение при прокрутке колесом мыши
            ScrollViewer.ScrollToVerticalOffset(ScrollViewer.VerticalOffset - e.Delta);
            // Пометьте событие как обработанное, чтобы оно не передалось другим элементам
            e.Handled = true;
        }

        /// <summary>
        /// Устанавливает таймер на автозакрытие
        /// </summary>
        internal void Start_WindowClose()
        {
            // Запускаем таймер при загрузке формы, чтобы окно провисело гарантировано 15 сек.
            Dispatcher.Invoke(() =>
            {
                // Запускаем таймер при загрузке формы, чтобы окно провисело гарантировано 15 сек.
                _closeTimer.Interval = 25000;
                _closeTimer.Start();
            });
        }

        private void Application_ModuleStatus(LoaderEvantEntity lModule, Brush brush)
        {
            Dispatcher.Invoke(() =>
            {
                lModule.LoadColor = brush;
                _loadModules.Add(lModule);

                modulesListBox.ItemsSource = _loadModules;
            });
        }

        /// <summary>
        /// Обработчик события RiseStepProgress
        /// </summary>
        private void Application_Progress(MainStatus mainStatus, string toolTip, Brush brush)
        {
            Dispatcher.Invoke(() =>
            {
                
                LoaderStatusEntity stEntity = _loaderStatusEntitys.Where(x => x.CurrentMainStatus == mainStatus).FirstOrDefault();
                if (stEntity != null)
                {
                    stEntity.StrStatus = _statusDone;
                    stEntity.CurrentToolTip = toolTip;
                    stEntity.CurrentStrStatusColor = brush;
                };
            });
        }

        private void CloseTimer_Tick(object sender, EventArgs e)
        {
            // Закрываем форму после истечения задержки.
            _closeTimer.Stop();
            Close();
        }

        private void Instruction_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            string loaderDescriptionURL = _loaderDescription.InstructionURL;
            if (loaderDescriptionURL != null)
                Process.Start(new ProcessStartInfo(loaderDescriptionURL) { UseShellExecute = true });
        }

        private void BtnLike_Click(object sender, RoutedEventArgs e)
        {
            SetUnclickableRateBtns();
            LikeStatus?.Invoke(MainDB_LoaderDescriptions_RateType.Approval, _loaderDescription);
        }

        private void BtnDislike_Click(object sender, RoutedEventArgs e)
        {
            SetUnclickableRateBtns();
            LikeStatus?.Invoke(MainDB_LoaderDescriptions_RateType.Disapproval, _loaderDescription);
        }

        /// <summary>
        /// Установить кнопки рейтинга не нажимаемыми
        /// </summary>
        private void SetUnclickableRateBtns()
        {
            btnLike.IsEnabled = false;
            btnLike.Foreground = Brushes.Gray;
            btnDislike.IsEnabled = false;
            btnDislike.Foreground = Brushes.Gray;
        }
    }
}
