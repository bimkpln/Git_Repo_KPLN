using System;
using System.Drawing;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_TaskManager.Forms
{
    /// <summary>
    /// Класс по созданию скрина. АРХИВ, пока оставь, если не зайдёт метод вставить скрин из буфера
    /// </summary>
    public partial class ImgSelectionFrom : Window
    {
        private System.Windows.Point startPoint;
        private System.Windows.Point endPoint;
        private bool isSelecting = false;

        public ImgSelectionFrom()
        {
            InitializeComponent();
            this.PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        public Bitmap CapturedImage { get; private set; }

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape) Close();
        }

        private void SelectionCanvas_MouseDown(object sender, MouseButtonEventArgs e)
        {
            startPoint = e.GetPosition(this);
            isSelecting = true;
            SelectionRectangle.Visibility = Visibility.Visible;
        }

        private void SelectionCanvas_MouseMove(object sender, MouseEventArgs e)
        {
            if (!isSelecting) return;

            endPoint = e.GetPosition(this);
            double x = Math.Min(startPoint.X, endPoint.X);
            double y = Math.Min(startPoint.Y, endPoint.Y);
            double width = Math.Abs(startPoint.X - endPoint.X);
            double height = Math.Abs(startPoint.Y - endPoint.Y);

            Canvas.SetLeft(SelectionRectangle, x);
            Canvas.SetTop(SelectionRectangle, y);
            SelectionRectangle.Width = width;
            SelectionRectangle.Height = height;
        }

        private void SelectionCanvas_MouseUp(object sender, MouseButtonEventArgs e)
        {
            isSelecting = false;
            CaptureSelectedArea();
            this.DialogResult = true;
        }

        private void CaptureSelectedArea()
        {
            var screenStart = PointToScreen(startPoint);
            var screenEnd = PointToScreen(endPoint);

            int x = (int)Math.Min(screenStart.X, screenEnd.X);
            int y = (int)Math.Min(screenStart.Y, screenEnd.Y);
            int width = (int)Math.Abs(screenStart.X - screenEnd.X);
            int height = (int)Math.Abs(screenStart.Y - screenEnd.Y);

            if (width > 0 && height > 0)
            {
                Bitmap bmp = new Bitmap(width, height);
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.CopyFromScreen(x, y, 0, 0, new System.Drawing.Size(width, height));
                }
                CapturedImage = bmp;
            }
        }
    }
}
