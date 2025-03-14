using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Interop;

namespace KPLN_TaskManager.Services
{
    internal static class ClipboardService
    {
        private const int WM_CLIPBOARDUPDATE = 0x031D;

        private static Window _window;

        private static IntPtr _nextClipboardViewer;

        [DllImport("user32.dll")]
        private static extern IntPtr SetClipboardViewer(IntPtr hWndNewViewer);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool ChangeClipboardChain(IntPtr hWndRemove, IntPtr hWndNewNext);

        internal static void StartClipboardListener(Window wind)
        {
            _window = wind;

            IntPtr hwnd = new WindowInteropHelper(_window).Handle;
            _nextClipboardViewer = SetClipboardViewer(hwnd);
            ComponentDispatcher.ThreadPreprocessMessage += OnClipboardUpdate;
        }

        private static void StopClipboardListener()
        {
            ComponentDispatcher.ThreadPreprocessMessage -= OnClipboardUpdate;
            ChangeClipboardChain(new WindowInteropHelper(_window).Handle, _nextClipboardViewer);
        }

        private static void OnClipboardUpdate(ref MSG msg, ref bool handled)
        {
            if (msg.message == WM_CLIPBOARDUPDATE)
            {
                if (System.Windows.Forms.Clipboard.ContainsImage())
                {
                    //byte[] imageData = GetClipboardImage();
                    //SaveScreenshotToDatabase(imageData, "Data Source=your_database.db;");
                    //MessageBox.Show("Скрыншот захаваны!");

                    //StopClipboardListener();
                }
                handled = true;
            }
        }
    }
}
