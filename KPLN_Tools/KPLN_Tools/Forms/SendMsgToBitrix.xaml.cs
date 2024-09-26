using KPLN_Tools.Forms.Models;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using static KPLN_Library_Forms.Common.UIStatus;

namespace KPLN_Tools.Forms
{
    public partial class SendMsgToBitrix : Window
    {

        public SendMsgToBitrix(ObservableCollection<SendMsgToBitrix_UserEntity> userEntitysCollection)
        {
            CurrentViewModel = new SendMsgToBitrix_ViewModel(userEntitysCollection);
            InitializeComponent();

            //// Устанавливаем DataContext для привязки данных
            DataContext = CurrentViewModel;

            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        /// <summary>
        /// Модель для окна wpf
        /// </summary>
        public SendMsgToBitrix_ViewModel CurrentViewModel { get; set; }

        /// <summary>
        /// Статус запуска
        /// </summary>
        public RunStatus Status { get; private set; } = RunStatus.Close;

        private static T FindVisualChild<T>(DependencyObject obj) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(obj); i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(obj, i);
                if (child != null && child is T)
                    return (T)child;
                else
                {
                    T childOfChild = FindVisualChild<T>(child);
                    if (childOfChild != null)
                        return childOfChild;
                }
            }

            return null;
        }

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Status = RunStatus.Close;
                Close();
            }

            if (e.Key == Key.Enter)
                RunBtn_Click(sender, e);
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            Keyboard.Focus(UserInput_Surname);

            // Подписываюсь на выбор элемента (IsChecked) в Users
            foreach (var item in Users.Items)
            {
                // Получаем контейнер для элемента
                if (Users.ItemContainerGenerator.ContainerFromItem(item) is FrameworkElement container)
                {
                    // Ищем CheckBox в визуальном дереве контейнера
                    var checkBox = FindVisualChild<CheckBox>(container);
                    if (checkBox != null)
                    {
                        checkBox.Checked += CheckBox_Checked;
                        checkBox.Unchecked += CheckBox_Unchecked;
                    }
                }
            }
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e) => UpdateRunButtonState();

        private void CheckBox_Unchecked(object sender, RoutedEventArgs e) => UpdateRunButtonState();

        private void UpdateRunButtonState()
        {
            // Проверим, есть ли выбранные CheckBox
            RunBtn.IsEnabled = Users.Items.Cast<object>()
                .Select(item => Users.ItemContainerGenerator.ContainerFromItem(item) as FrameworkElement)
                .Select(container => FindVisualChild<CheckBox>(container))
                .Any(cb => cb != null && cb.IsChecked == true);
        }

        private void ScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            // Выполните здесь нужное вам поведение при прокрутке колесом мыши
            UsersScroll.ScrollToVerticalOffset(UsersScroll.VerticalOffset - e.Delta);
            // Пометьте событие как обработанное, чтобы оно не передалось другим элементам
            e.Handled = true;
        }

        private void RunBtn_Click(object sender, RoutedEventArgs e)
        {
            if (CurrentViewModel.SelectedElements.Any())
            {
                Status = RunStatus.Run;
                this.Close();
            }
        }
    }
}
