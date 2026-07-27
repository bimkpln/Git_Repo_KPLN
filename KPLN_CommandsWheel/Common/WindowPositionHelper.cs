using System;
using System.Windows;
using System.Windows.Interop;
using WinForms = System.Windows.Forms;

namespace KPLN_CommandsWheel.Services
{
    public static class WindowPositionHelper
    {
        public static void ShowNearCursor(Window window, double offsetX = 16, double offsetY = 16)
        {
            window.WindowStartupLocation = WindowStartupLocation.Manual;

            window.SourceInitialized += (sender, args) =>
            {
                MoveNearCursor(window, offsetX, offsetY);
            };
        }

        public static void ShowCenteredOnCursor(Window window)
        {
            window.WindowStartupLocation = WindowStartupLocation.Manual;

            window.SourceInitialized += (sender, args) =>
            {
                MoveCenteredOnCursor(window);
            };
        }

        private static void MoveNearCursor(Window window, double offsetX, double offsetY)
        {
            var cursorPosition = WinForms.Cursor.Position;

            PresentationSource source = PresentationSource.FromVisual(window);

            if (source?.CompositionTarget == null)
            {
                window.Left = cursorPosition.X + offsetX;
                window.Top = cursorPosition.Y + offsetY;
                return;
            }

            Point cursorPoint = source.CompositionTarget.TransformFromDevice.Transform(
                new Point(cursorPosition.X, cursorPosition.Y)
            );

            window.Left = cursorPoint.X + offsetX;
            window.Top = cursorPoint.Y + offsetY;
        }

        private static void MoveCenteredOnCursor(Window window)
        {
            var cursorPosition = WinForms.Cursor.Position;
            double width = window.ActualWidth > 0 ? window.ActualWidth : window.Width;
            double height = window.ActualHeight > 0 ? window.ActualHeight : window.Height;

            PresentationSource source = PresentationSource.FromVisual(window);

            if (source?.CompositionTarget == null)
            {
                window.Left = cursorPosition.X - width / 2;
                window.Top = cursorPosition.Y - height / 2;
                return;
            }

            Point cursorPoint = source.CompositionTarget.TransformFromDevice.Transform(
                new Point(cursorPosition.X, cursorPosition.Y)
            );

            window.Left = cursorPoint.X - width / 2;
            window.Top = cursorPoint.Y - height / 2;
        }
    }
}
