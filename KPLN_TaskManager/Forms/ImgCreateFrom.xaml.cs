using System.Drawing;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace KPLN_TaskManager.Forms
{
    /// <summary>
    /// Класс по созданию окна работы со скрином. АРХИВ, пока оставь, если не зайдёт метод вставить скрин из буфера
    /// </summary>
    public partial class ImgCreateFrom : Window
    {
        public ImgCreateFrom(Window owner)
        {
            InitializeComponent();
            inkCanvas.EditingMode = InkCanvasEditingMode.None;
            Owner = owner;

            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        public byte[] ResultImageData { get; private set; }

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape) Close();
        }

        private void BtnCapture_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();

            ImgSelectionFrom imgSelectionFrom = new ImgSelectionFrom(this);
            bool? formResult = imgSelectionFrom.ShowDialog();

            if ((bool)formResult)
            {
                Bitmap screenshot = imgSelectionFrom.CapturedImage;
                inkCanvas.Background = new ImageBrush(BitmapToImageSource(screenshot));

            }

            this.Show();
        }

        private void EnableMarkerMode_Click(object sender, RoutedEventArgs e)
        {
            inkCanvas.EditingMode = InkCanvasEditingMode.Ink;
            inkCanvas.DefaultDrawingAttributes = new DrawingAttributes
            {
                Color = Colors.Red,
                Width = 3,
                Height = 3
            };
        }

        private void EnableTextMode_Click(object sender, RoutedEventArgs e)
        {

        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            ResultImageData = SaveToByteArray(inkCanvas);
        }

        private void InkCanvas_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            System.Windows.Point position = e.GetPosition(inkCanvas);
            AddTextBoxAtPosition(position);
        }

        private void AddTextBoxAtPosition(System.Windows.Point position)
        {
            TextBox textBox = new TextBox
            {
                Width = 200,
                Height = 30,
                Background = System.Windows.Media.Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Foreground = System.Windows.Media.Brushes.Black
            };
            Canvas.SetLeft(textBox, position.X);
            Canvas.SetTop(textBox, position.Y);
            inkCanvas.Children.Add(textBox);
            textBox.Focus();
        }

        private BitmapImage BitmapToImageSource(Bitmap bitmap)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                ms.Position = 0;
                BitmapImage bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = ms;
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();
                return bitmapImage;
            }
        }

        private byte[] SaveToByteArray(InkCanvas inkCanvas)
        {
            RenderTargetBitmap rtb = new RenderTargetBitmap(
                (int)inkCanvas.ActualWidth, (int)inkCanvas.ActualHeight, 96, 96, PixelFormats.Pbgra32);
            rtb.Render(inkCanvas);

            PngBitmapEncoder encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(rtb));

            using (MemoryStream ms = new MemoryStream())
            {
                encoder.Save(ms);
                return ms.ToArray();
            }
        }
    }
}
