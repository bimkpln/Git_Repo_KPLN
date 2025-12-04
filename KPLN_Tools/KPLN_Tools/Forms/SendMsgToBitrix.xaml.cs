using KPLN_Tools.Forms.Models;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using static KPLN_Library_Forms.Common.UIStatus;

namespace KPLN_Tools.Forms
{
    public partial class SendMsgToBitrix : Window
    {
        public SendMsgToBitrix(ObservableCollection<SendMsgToBitrix_UserEntity> userEntitysCollection)
        {
            CurrentViewModel = new SendMsgToBitrix_VM(userEntitysCollection);
            InitializeComponent();

            //// Устанавливаем DataContext для привязки данных
            DataContext = CurrentViewModel;

            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        /// <summary>
        /// Модель для окна wpf
        /// </summary>
        public SendMsgToBitrix_VM CurrentViewModel { get; set; }

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
                Close();
            else if (e.Key == Key.Enter)
                RunBtn_Click(sender, e);
        }

        private void OnLoaded(object sender, RoutedEventArgs e) => Keyboard.Focus(UserInput_Surname);

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
                DialogResult = true;
                this.Close();
            }
            else
                System.Windows.Forms.MessageBox.Show("Сначала выбери адресат/-ы для отправки сообщения", "Внимание", System.Windows.Forms.MessageBoxButtons.OK);
        }
    }
}
