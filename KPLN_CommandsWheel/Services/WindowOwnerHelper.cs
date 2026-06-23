using System;
using System.Windows;
using System.Windows.Interop;

namespace KPLN_CommandsWheel.Services
{
    internal static class WindowOwnerHelper
    {
        internal static void Apply(Window window)
        {
            if (window == null || ModuleData.RevitMainWindowHandle == IntPtr.Zero)
            {
                return;
            }

            WindowInteropHelper helper = new WindowInteropHelper(window);
            helper.Owner = ModuleData.RevitMainWindowHandle;
            window.WindowStartupLocation = WindowStartupLocation.CenterOwner;
        }
    }
}