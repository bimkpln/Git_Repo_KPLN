using KPLN_Library_Forms.Common;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using static KPLN_Library_Forms.Common.UIStatus;

namespace KPLN_Library_Forms.UI
{
    public partial class ButtonToRun : Window
    {
        /// <summary>
        /// Флаг для идентификации запуска приложения, а не закрытия через Х (любое закрытие окна связано с Window_Closing, поэтому нужен доп. флаг)
        /// </summary>
        private bool _isRun = false;
        private readonly string _mainDescription = string.Empty;

        private readonly IEnumerable<ButtonToRunEntity> _collection;

        public ButtonToRun(string mainTitle, IEnumerable<ButtonToRunEntity> collection)
        {
            _mainDescription = mainTitle;
            _collection = collection;
            InitializeComponent();

            BtnColl.ItemsSource = _collection;
            MainDescription.Text = _mainDescription;
            PreviewKeyDown += new KeyEventHandler(HandleEsc);
        }

        /// <summary>
        /// Коллекция элементов, которые будут отображаться в окне
        /// </summary>
        public IEnumerable<ButtonToRunEntity> Collection { get { return _collection; } }

        /// <summary>
        /// Нажатая кнопка
        /// </summary>
        public ButtonToRunEntity SelectedButton { get; private set; }

        /// <summary>
        /// Статус запуска
        /// </summary>
        public RunStatus Status { get; private set; }

        private void HandleEsc(object sender, KeyEventArgs e)
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

        private void OnBtnClick(object sender, RoutedEventArgs e)
        {
            SelectedButton = (sender as Button).DataContext as ButtonToRunEntity;

            _isRun = true;
            Status = RunStatus.Run;
            Close();
        }
    }
}
