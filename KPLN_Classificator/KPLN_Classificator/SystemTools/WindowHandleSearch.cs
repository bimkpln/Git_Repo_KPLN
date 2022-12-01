using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;

namespace KPLN_Classificator.SystemTools
{
    public class WindowHandleSearch : IWin32Window, System.Windows.Forms.IWin32Window
    {
        #region Static methods
        static public WindowHandleSearch MainWindowHandle
        {
            get
            {
                var revitProcess = Process.GetCurrentProcess();
                return new WindowHandleSearch(GetMainWindow(revitProcess.Id));
            }
        }
        #endregion

        #region Constructor
        public WindowHandleSearch(IntPtr hwnd)
        {
            Debug.Assert(IntPtr.Zero != hwnd,
              "Null window handle");

            Handle = hwnd;
        }
        #endregion
        #region Methods
        public IntPtr Handle { get; private set; }
        public void SetAsOwner(Window childWindow)
        {
            var helper = new WindowInteropHelper(childWindow)
            { Owner = Handle };
        }

        private delegate bool EnumWindowsProc(IntPtr hWnd, int lParam);

        [DllImport("user32.DLL")]
        private static extern bool EnumWindows(EnumWindowsProc enumFunc, int lParam);

        [DllImport("user32.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
        private static extern IntPtr GetParent(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.DLL")]
        private static extern IntPtr GetShellWindow();

        [DllImport("user32.DLL")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.DLL")]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.DLL")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        public static IntPtr GetMainWindow(int pid)
        {
            IntPtr shellWindow = GetShellWindow();
            List<IntPtr> windowsForPid = new List<IntPtr>();

            try
            {
                EnumWindows(
                delegate (IntPtr hWnd, int lParam)
                {
                    if (hWnd == shellWindow) return true;
                    if (!IsWindowVisible(hWnd)) return true;

                    uint windowPid = 0;
                    GetWindowThreadProcessId(hWnd, out windowPid);

                    if (windowPid == pid)
                    {
                        IntPtr parentHwnd = GetParent(hWnd);
                        if (parentHwnd == IntPtr.Zero)
                            windowsForPid.Add(hWnd);
                    }

                    return true;
                }
                , 0);
            }
            catch (Exception)
            { }

            return DetermineMainWindow(windowsForPid);
        }
        private static IntPtr DetermineMainWindow(
          List<IntPtr> handles)
        {
            if (handles == null || handles.Count <= 0)
                return IntPtr.Zero;
            IntPtr mainWindow = IntPtr.Zero;
            if (handles.Count == 1)
            {
                mainWindow = handles[0];
            }
            else
            {
                foreach (var hWnd in handles)
                {
                    int length = GetWindowTextLength(hWnd);
                    if (length == 0) continue;

                    StringBuilder builder = new StringBuilder(
                      length);

                    GetWindowText(hWnd, builder, length + 1);
                    if (builder.ToString().ToLower().Contains(
                      "autodesk revit"))
                    {
                        mainWindow = hWnd;
                        break;
                    }
                }
            }
            return mainWindow;
        }
        #endregion
    }
}
