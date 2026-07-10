using Autodesk.Revit.UI;
using KPLN_CommandsWheel.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Input;
using System.Windows.Interop;

namespace KPLN_CommandsWheel.Services
{
    internal static class HotkeyService
    {
        private const int WH_KEYBOARD_LL = 13;
        private const int WH_MOUSE_LL = 14;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;
        private const int WM_SYSKEYDOWN = 0x0104;
        private const int WM_SYSKEYUP = 0x0105;
        private const int WM_XBUTTONDOWN = 0x020B;
        private const int WM_HOTKEY = 0x0312;
        private const int XBUTTON1 = 1;
        private const int XBUTTON2 = 2;
        private const int CommandSearchHotkeyId = 0x4B50;
        private const int CommandsWheelHotkeyId = 0x4B51;
        private const uint MOD_ALT = 0x0001;
        private const uint MOD_CONTROL = 0x0002;
        private const uint MOD_SHIFT = 0x0004;
        private const uint MOD_NOREPEAT = 0x4000;

        private static readonly HashSet<string> PressedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static HookProc _keyboardProc;
        private static HookProc _mouseProc;
        private static IntPtr _keyboardHook = IntPtr.Zero;
        private static IntPtr _mouseHook = IntPtr.Zero;
        private static ExternalEvent _openCommandSearchEvent;
        private static ExternalEvent _openCommandsWheelEvent;
        private static HwndSource _hotkeySource;
        private static UserSettings _settings;
        private static bool _isInitialized;
        private static bool _isSuspended;
        private static bool _isCommandSearchHotkeyRegistered;
        private static bool _isCommandsWheelHotkeyRegistered;
        private static bool _searchHotkeyActive;
        private static bool _wheelHotkeyActive;
        private static HotkeyTarget? _lastRaisedTarget;
        private static DateTime _lastRaiseTimeUtc;

        internal static void Initialize()
        {
            ReloadSettings();
            EnsureExternalEvents();

            if (_isInitialized)
            {
                RegisterKeyboardHotkeys();
                return;
            }

            _keyboardProc = KeyboardHookCallback;
            _mouseProc = MouseHookCallback;

            EnsureHotkeyWindow();
            RegisterKeyboardHotkeys();

            _keyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, _keyboardProc, IntPtr.Zero, 0);
            if (_keyboardHook == IntPtr.Zero)
            {
                _keyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, _keyboardProc, GetCurrentProcessModuleHandle(), 0);
            }

            _mouseHook = SetWindowsHookEx(WH_MOUSE_LL, _mouseProc, IntPtr.Zero, 0);
            if (_mouseHook == IntPtr.Zero)
            {
                _mouseHook = SetWindowsHookEx(WH_MOUSE_LL, _mouseProc, GetCurrentProcessModuleHandle(), 0);
            }

            _isInitialized = true;
        }

        internal static void ReloadSettings(UserSettings settings = null)
        {
            _settings = settings ?? UserSettingsService.Load();
            PressedKeys.Clear();
            _searchHotkeyActive = false;
            _wheelHotkeyActive = false;

            if (_isInitialized)
            {
                RegisterKeyboardHotkeys();
            }
        }

        internal static void Shutdown()
        {
            UnregisterKeyboardHotkeys();

            if (_keyboardHook != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_keyboardHook);
                _keyboardHook = IntPtr.Zero;
            }

            if (_mouseHook != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_mouseHook);
                _mouseHook = IntPtr.Zero;
            }

            if (_hotkeySource != null)
            {
                _hotkeySource.RemoveHook(HotkeyWindowHook);
                _hotkeySource.Dispose();
                _hotkeySource = null;
            }

            _isInitialized = false;
            _isSuspended = false;
            PressedKeys.Clear();
        }

        internal static void SuspendHotkeys()
        {
            _isSuspended = true;
            PressedKeys.Clear();
            _searchHotkeyActive = false;
            _wheelHotkeyActive = false;
            UnregisterKeyboardHotkeys();
        }

        internal static void ResumeHotkeys()
        {
            _isSuspended = false;
            PressedKeys.Clear();
            _searchHotkeyActive = false;
            _wheelHotkeyActive = false;

            if (_isInitialized)
            {
                RegisterKeyboardHotkeys();
            }
        }

        private static void EnsureExternalEvents()
        {
            try
            {
                if (_openCommandSearchEvent == null)
                {
                    _openCommandSearchEvent = ExternalEvent.Create(new OpenWindowRequestHandler(HotkeyTarget.CommandSearch));
                }

                if (_openCommandsWheelEvent == null)
                {
                    _openCommandsWheelEvent = ExternalEvent.Create(new OpenWindowRequestHandler(HotkeyTarget.CommandsWheel));
                }
            }
            catch
            {
                return;
            }
        }

        private static void EnsureHotkeyWindow()
        {
            if (_hotkeySource != null)
            {
                return;
            }

            HwndSourceParameters parameters = new HwndSourceParameters("KPLN Commands Wheel Hotkeys")
            {
                Width = 0,
                Height = 0,
                WindowStyle = unchecked((int)0x80000000)
            };

            _hotkeySource = new HwndSource(parameters);
            _hotkeySource.AddHook(HotkeyWindowHook);
        }

        private static void RegisterKeyboardHotkeys()
        {
            EnsureHotkeyWindow();
            UnregisterKeyboardHotkeys();

            if (_isSuspended || _hotkeySource == null || _hotkeySource.Handle == IntPtr.Zero || _settings == null)
            {
                return;
            }

            _isCommandSearchHotkeyRegistered = TryRegisterKeyboardHotkey(
                CommandSearchHotkeyId,
                _settings.CommandSearchHotkey);

            _isCommandsWheelHotkeyRegistered = TryRegisterKeyboardHotkey(
                CommandsWheelHotkeyId,
                _settings.CommandsWheelHotkey);
        }

        private static void UnregisterKeyboardHotkeys()
        {
            if (_hotkeySource == null || _hotkeySource.Handle == IntPtr.Zero)
            {
                _isCommandSearchHotkeyRegistered = false;
                _isCommandsWheelHotkeyRegistered = false;
                return;
            }

            if (_isCommandSearchHotkeyRegistered)
            {
                UnregisterHotKey(_hotkeySource.Handle, CommandSearchHotkeyId);
                _isCommandSearchHotkeyRegistered = false;
            }

            if (_isCommandsWheelHotkeyRegistered)
            {
                UnregisterHotKey(_hotkeySource.Handle, CommandsWheelHotkeyId);
                _isCommandsWheelHotkeyRegistered = false;
            }
        }

        private static bool TryRegisterKeyboardHotkey(int id, HotkeyGesture gesture)
        {
            uint modifiers;
            uint virtualKey;
            if (!TryGetRegisterHotKeyParts(gesture, out modifiers, out virtualKey))
            {
                return false;
            }

            return RegisterHotKey(_hotkeySource.Handle, id, modifiers | MOD_NOREPEAT, virtualKey);
        }

        private static bool TryGetRegisterHotKeyParts(HotkeyGesture gesture, out uint modifiers, out uint virtualKey)
        {
            modifiers = 0;
            virtualKey = 0;

            if (HotkeyGestureService.IsEmpty(gesture) || !string.IsNullOrWhiteSpace(gesture.MouseButton))
            {
                return false;
            }

            List<string> keys = HotkeyGestureService.NormalizeKeys(gesture.Keys);
            List<string> mainKeys = keys
                .Where(key => !HotkeyGestureService.IsModifier(key))
                .ToList();

            if (mainKeys.Count != 1)
            {
                return false;
            }

            foreach (string key in keys)
            {
                if (string.Equals(key, "Ctrl", StringComparison.OrdinalIgnoreCase))
                {
                    modifiers |= MOD_CONTROL;
                }
                else if (string.Equals(key, "Shift", StringComparison.OrdinalIgnoreCase))
                {
                    modifiers |= MOD_SHIFT;
                }
                else if (string.Equals(key, "Alt", StringComparison.OrdinalIgnoreCase))
                {
                    modifiers |= MOD_ALT;
                }
            }

            if (modifiers == 0)
            {
                return false;
            }

            Key parsedKey;
            if (!TryParseKey(mainKeys[0], out parsedKey))
            {
                return false;
            }

            int keyValue = KeyInterop.VirtualKeyFromKey(parsedKey);
            if (keyValue == 0)
            {
                return false;
            }

            virtualKey = (uint)keyValue;
            return true;
        }

        private static bool TryParseKey(string value, out Key key)
        {
            if (string.Equals(value, "Enter", StringComparison.OrdinalIgnoreCase))
            {
                key = Key.Return;
                return true;
            }

            return Enum.TryParse(value, true, out key);
        }

        private static IntPtr HotkeyWindowHook(IntPtr hwnd, int message, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (_isSuspended || message != WM_HOTKEY)
            {
                return IntPtr.Zero;
            }

            int id = wParam.ToInt32();
            if (id == CommandSearchHotkeyId)
            {
                handled = true;
                if (IsForegroundCurrentProcess())
                {
                    Raise(HotkeyTarget.CommandSearch);
                }
            }
            else if (id == CommandsWheelHotkeyId)
            {
                handled = true;
                if (IsForegroundCurrentProcess())
                {
                    Raise(HotkeyTarget.CommandsWheel);
                }
            }

            return IntPtr.Zero;
        }

        private static IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (!_isSuspended && nCode >= 0 && IsForegroundCurrentProcess())
            {
                int message = wParam.ToInt32();
                Kbdllhookstruct keyboardData = (Kbdllhookstruct)Marshal.PtrToStructure(lParam, typeof(Kbdllhookstruct));
                string keyName = GetKeyNameFromVirtualKey(keyboardData.vkCode);

                if (!string.IsNullOrWhiteSpace(keyName))
                {
                    if (message == WM_KEYDOWN || message == WM_SYSKEYDOWN)
                    {
                        PressedKeys.Add(keyName);
                        if (TryTriggerKeyboardHotkey())
                        {
                            return new IntPtr(1);
                        }
                    }
                    else if (message == WM_KEYUP || message == WM_SYSKEYUP)
                    {
                        PressedKeys.Remove(keyName);
                        RefreshKeyboardHotkeyActiveState();
                    }
                }
            }

            return CallNextHookEx(_keyboardHook, nCode, wParam, lParam);
        }

        private static IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (!_isSuspended && nCode >= 0 && wParam.ToInt32() == WM_XBUTTONDOWN && IsForegroundCurrentProcess())
            {
                Msllhookstruct mouseData = (Msllhookstruct)Marshal.PtrToStructure(lParam, typeof(Msllhookstruct));
                int xButton = (mouseData.mouseData >> 16) & 0xffff;
                string mouseButton = xButton == XBUTTON1 ? "XButton1" : xButton == XBUTTON2 ? "XButton2" : null;

                if (!string.IsNullOrWhiteSpace(mouseButton) && TryTriggerMouseHotkey(mouseButton))
                {
                    return new IntPtr(1);
                }
            }

            return CallNextHookEx(_mouseHook, nCode, wParam, lParam);
        }

        private static bool TryTriggerKeyboardHotkey()
        {
            if (_settings == null)
            {
                ReloadSettings();
            }

            bool searchMatches = HotkeyGestureService.Matches(_settings.CommandSearchHotkey, PressedKeys, null);
            bool wheelMatches = HotkeyGestureService.Matches(_settings.CommandsWheelHotkey, PressedKeys, null);

            if (searchMatches)
            {
                if (!_searchHotkeyActive)
                {
                    Raise(HotkeyTarget.CommandSearch);
                }

                _searchHotkeyActive = true;
                return true;
            }

            if (wheelMatches)
            {
                if (!_wheelHotkeyActive)
                {
                    Raise(HotkeyTarget.CommandsWheel);
                }

                _wheelHotkeyActive = true;
                return true;
            }

            RefreshKeyboardHotkeyActiveState();
            return false;
        }

        private static bool TryTriggerMouseHotkey(string mouseButton)
        {
            if (_settings == null)
            {
                ReloadSettings();
            }

            if (HotkeyGestureService.Matches(_settings.CommandSearchHotkey, PressedKeys, mouseButton))
            {
                Raise(HotkeyTarget.CommandSearch);
                return true;
            }

            if (HotkeyGestureService.Matches(_settings.CommandsWheelHotkey, PressedKeys, mouseButton))
            {
                Raise(HotkeyTarget.CommandsWheel);
                return true;
            }

            return false;
        }

        private static void RefreshKeyboardHotkeyActiveState()
        {
            if (_settings == null)
            {
                _searchHotkeyActive = false;
                _wheelHotkeyActive = false;
                return;
            }

            _searchHotkeyActive = HotkeyGestureService.Matches(_settings.CommandSearchHotkey, PressedKeys, null);
            _wheelHotkeyActive = HotkeyGestureService.Matches(_settings.CommandsWheelHotkey, PressedKeys, null);
        }

        private static bool Raise(HotkeyTarget target)
        {
            DateTime now = DateTime.UtcNow;
            if (_lastRaisedTarget.HasValue
                && _lastRaisedTarget.Value == target
                && (now - _lastRaiseTimeUtc).TotalMilliseconds < 250)
            {
                return false;
            }

            _lastRaisedTarget = target;
            _lastRaiseTimeUtc = now;
            EnsureExternalEvents();

            try
            {
                if (target == HotkeyTarget.CommandSearch)
                {
                    if (_openCommandSearchEvent == null)
                    {
                        return false;
                    }

                    _openCommandSearchEvent.Raise();
                    return true;
                }

                if (_openCommandsWheelEvent == null)
                {
                    return false;
                }

                _openCommandsWheelEvent.Raise();
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string GetKeyNameFromVirtualKey(int virtualKey)
        {
            switch (virtualKey)
            {
                case 0x10:
                case 0xA0:
                case 0xA1:
                    return "Shift";

                case 0x11:
                case 0xA2:
                case 0xA3:
                    return "Ctrl";

                case 0x12:
                case 0xA4:
                case 0xA5:
                    return "Alt";

                default:
                    return HotkeyGestureService.NormalizeKey(KeyInterop.KeyFromVirtualKey(virtualKey));
            }
        }

        private static bool IsForegroundCurrentProcess()
        {
            IntPtr foregroundWindow = GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero)
            {
                return false;
            }

            int processId;
            GetWindowThreadProcessId(foregroundWindow, out processId);
            return processId == Process.GetCurrentProcess().Id;
        }

        private static IntPtr GetCurrentProcessModuleHandle()
        {
            using (Process process = Process.GetCurrentProcess())
            {
                using (ProcessModule module = process.MainModule)
                {
                    return GetModuleHandle(module.ModuleName);
                }
            }
        }

        private enum HotkeyTarget
        {
            CommandSearch,
            CommandsWheel
        }

        private class OpenWindowRequestHandler : IExternalEventHandler
        {
            private readonly HotkeyTarget _target;

            internal OpenWindowRequestHandler(HotkeyTarget target)
            {
                _target = target;
            }

            public void Execute(UIApplication app)
            {
                if (_target == HotkeyTarget.CommandSearch)
                {
                    CommandWindowService.ShowCommandSearch(app);
                    return;
                }

                CommandWindowService.ShowCommandsWheel(app);
            }

            public string GetName()
            {
                return _target == HotkeyTarget.CommandSearch
                    ? "KPLN Commands Hotkey"
                    : "KPLN Commands Wheel Hotkey";
            }
        }

        private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct Kbdllhookstruct
        {
            public int vkCode;
            public int scanCode;
            public int flags;
            public int time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct Point
        {
            public int x;
            public int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct Msllhookstruct
        {
            public Point pt;
            public int mouseData;
            public int flags;
            public int time;
            public IntPtr dwExtraInfo;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int processId);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);
    }
}