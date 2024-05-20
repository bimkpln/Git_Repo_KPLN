using System.Windows;
using System.Windows.Input;
using static KPLN_Library_Forms.Common.UIStatus;

namespace KPLN_Library_Forms.UI
{
    /// <summary>
    /// Окно для вывода сообщения пользователю
    /// </summary>
    public partial class CustomMessageBox : Window
    {
        /// <summary>
        /// Флаг для идентификации запуска приложения, а не закрытия через Х (любое закрытие окна связано с Window_Closing, поэтому нужен доп. флаг)
        /// </summary>
        private bool _isRun = false;

        public CustomMessageBox(string tbHeader, string tblMainContent)
        {
            InitializeComponent();

            this.LabelHeader.Content = tbHeader;
            this.TblMainContent.Text = tblMainContent;

            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        /// <summary>
        /// Статус запуска
        /// </summary>
        public RunStatus Status { get; private set; }

        /// <summary>
        /// Показать окно с учетом родительского
        /// </summary>
        public void ShowByOwner(Window owner)
        {
            this.Owner = owner;
            this.SetPosition(owner);
            this.Show();
        }

        /// <summary>
        /// Позиционируем окно рядом с главным окном
        /// </summary>
        private void SetPosition(Window owner)
        {
            this.Left = owner.Left + owner.Width;
            this.Top = owner.Top;
        }

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Status = RunStatus.Close;
                Close();
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!_isRun)
            {
                Status = RunStatus.Close;
            }
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            _isRun = true;
            Status = RunStatus.Run;

            Close();
        }
    }
}
