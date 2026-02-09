using Autodesk.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_ViewsAndLists_Ribbon.Common
{
    internal static class RevitWindowUtil
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int maxLength);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        public static string GetRevitWindowTitle()
        {
            IntPtr hwnd = ComponentManager.ApplicationWindow;
            if (hwnd == IntPtr.Zero)
                return string.Empty;

            int length = GetWindowTextLength(hwnd);
            var sb = new StringBuilder(length + 1);

            GetWindowText(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }
    }
}
