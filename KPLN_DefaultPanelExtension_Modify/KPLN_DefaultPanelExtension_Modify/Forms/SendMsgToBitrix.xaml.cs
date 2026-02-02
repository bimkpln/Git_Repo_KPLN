using KPLN_DefaultPanelExtension_Modify.Forms.Models;
using KPLN_Library_Forms.Services;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace KPLN_DefaultPanelExtension_Modify.Forms
{
    public partial class SendMsgToBitrix : Window
    {
        public SendMsgToBitrix(ObservableCollection<SendMsgToBitrix_UserEntity> userEntitysCollection)
        {
            CurrentViewModel = new SendMsgToBitrix_VM(userEntitysCollection);
            InitializeComponent();

            DataContext = CurrentViewModel;

            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        /// <summary>
        /// Модель для окна wpf
        /// </summary>
        public SendMsgToBitrix_VM CurrentViewModel { get; set; }

        public void ImgDelete_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult astericsResult = MessageBox.Show(
                $"Удалить текущий рисунок?",
                "Внимание",
                MessageBoxButton.YesNo,
                MessageBoxImage.Asterisk);

            if (astericsResult == MessageBoxResult.No)
                return;

            CurrentViewModel.ImageBuffer = null;
            SetImgExpander();

            Button btn = (Button)sender;
            if (btn.DataContext is ImgLargeFrom imgForm)
                imgForm.Close();
        }

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

        /// <summary>
        /// Загрузка изображения из буфера
        /// </summary>
        private void ImgLoad_Click(object sender, RoutedEventArgs e)
        {
            if (Clipboard.ContainsImage())
            {
                BitmapSource image = Clipboard.GetImage();
                if (image != null)
                {
                    SetImageBuffer(image);
                    SetImgExpander();

                    MessageBox.Show(
                        "Рисунок успешно добавлен!",
                        "Загрузка рисунка",
                        MessageBoxButton.OK,
                        MessageBoxImage.Asterisk);

                }
            }
            else
            {
                MessageBox.Show(
                    "Создай рисунок любым удобным способом, скопируй его в буфер обмена, а потом нажми \"Загрузить изображение\"",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void BodyImg_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            ImgLargeFrom imgLargeFrom = new ImgLargeFrom(this, CurrentViewModel);
            WindowHandleSearch.MainWindowHandle.SetAsOwner(imgLargeFrom);

            imgLargeFrom.Show();
        }

        /// <summary>
        /// Установить ImageBuffer
        /// </summary>
        private void SetImageBuffer(BitmapSource image)
        {
            byte[] resultBit;
            JpegBitmapEncoder encoder = new JpegBitmapEncoder
            {
                QualityLevel = 100
            };

            using (MemoryStream stream = new MemoryStream())
            {
                encoder.Frames.Add(BitmapFrame.Create(image));
                encoder.Save(stream);
                resultBit = stream.ToArray();
                stream.Close();
            }

            CurrentViewModel.ImageBuffer = resultBit;
        }

        /// <summary>
        /// Настройка заголовка и "свёрнутости" Expander с изображением
        /// </summary>
        private void SetImgExpander()
        {
            if (CurrentViewModel.ImageBuffer != null)
            {
                ImgExpander.IsEnabled = true;
                ImgExpander.IsExpanded = true;
            }
            else
            {
                ImgExpander.IsEnabled = false;
                ImgExpander.IsExpanded = false;
            }
        }
    }
}
