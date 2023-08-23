using KPLN_Loader.Core.SQLiteData;
using KPLN_Loader.Forms.Common;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows;

namespace KPLN_Loader.Forms
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class LoaderStatusForm : Window
    {
        private readonly string _statusError = "❌";
        private readonly string _statusDone = "✔️";
        private readonly IEnumerable<LoaderStatusEntity> _loaderStatusEntitys;
        private readonly List<LoaderEvantEntity> _loadModules;
        private readonly System.Windows.Forms.Timer _closeTimer;
        private string _loaderDescriptionURL;

        /// <summary>
        /// Окно с демонстрацией прогресса
        /// </summary>
        internal LoaderStatusForm(Application application)
        {
            application.Progress += Application_Progress;
            application.LoadStatus += Application_ModuleStatus;

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

            Show();
        }

        /// <summary>
        /// Текстовое описание
        /// </summary>
        internal void SetInstruction(LoaderDescription loaderDescription)
        {
            tblInstruction.Text = loaderDescription.Description;
            _loaderDescriptionURL = loaderDescription.InstructionURL;
            if (_loaderDescriptionURL != null)
            {
                tblInstruction.TextDecorations = TextDecorations.Underline;
            }
        }

        /// <summary>
        /// Добавляет маркер при дебаге модулей
        /// </summary>
        /// <param name="user"></param>
        internal void CheckAndSetDebugStatusByUser(User user)
        {
            if (user.IsDebugMode)
                DebugModeTxt.Visibility = Visibility.Visible;
        }

        /// <summary>
        /// Устанавливает таймер на автозакрытие
        /// </summary>
        internal void Start_WindowClose()
        {
            // Запускаем таймер при загрузке формы, чтобы окно провисело гарантировано 15 сек.
            _closeTimer.Interval = 15000;
            _closeTimer.Start();
        }

        private void Application_ModuleStatus(LoaderEvantEntity lModule, System.Windows.Media.Brush brush)
        {
            lModule.LoadColor = brush;
            _loadModules.Add(lModule);

            modulesListBox.ItemsSource = _loadModules;
        }

        /// <summary>
        /// Обработчик события RiseStepProgress
        /// </summary>
        private void Application_Progress(MainStatus mainStatus, string toolTip, System.Windows.Media.Brush brush)
        {
            LoaderStatusEntity stEntity = _loaderStatusEntitys.Where(x => x.CurrentMainStatus == mainStatus).FirstOrDefault();
            if (stEntity != null)
            {
                stEntity.StrStatus = _statusDone;
                stEntity.CurrentToolTip = toolTip;
                stEntity.CurrentStrStatusColor = brush;
            };
        }

        private void CloseTimer_Tick(object sender, EventArgs e)
        {
            // Закрываем форму после истечения задержки.
            _closeTimer.Stop();
            Close();
        }

        private void Instruction_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (_loaderDescriptionURL != null)
                Process.Start(new ProcessStartInfo(_loaderDescriptionURL) { UseShellExecute = true });
        }
    }
}
