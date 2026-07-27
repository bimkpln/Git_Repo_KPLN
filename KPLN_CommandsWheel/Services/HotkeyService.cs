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
        private const int WM_HOTKEY = 0x0312;
        private const int CommandSearchHotkeyId = 0x4B50;
        private const int CommandsWheelHotkeyId = 0x4B51;
        private const uint MOD_ALT = 0x0001;
        private const uint MOD_CONTROL = 0x0002;
        private const uint MOD_SHIFT = 0x0004;
        private const uint MOD_NOREPEAT = 0x4000;

        private static ExternalEvent _openCommandSearchEvent;
        private static ExternalEvent _openCommandsWheelEvent;
        private static HwndSource _hotkeySource;
        private static UserSettings _settings;
        private static bool _isInitialized;
        private static bool _isSuspended;
        private static bool _isCommandSearchHotkeyRegistered;
        private static bool _isCommandsWheelHotkeyRegistered;
        private static HotkeyTarget? _lastRaisedTarget;
        private static DateTime _lastRaiseTimeUtc;
        internal static void Initialize()
        {
            if (_isInitialized)
            {
                EnsureExternalEvents();
                return;
            }

            ReloadSettings();
            EnsureExternalEvents();

            EnsureHotkeyWindow();
            _isInitialized = true;
            RegisterKeyboardHotkeys();
        }

        internal static void ReloadSettings(UserSettings settings = null)
        {
            _settings = settings ?? UserSettingsService.Load();

            if (_isInitialized)
            {
                RegisterKeyboardHotkeys();
            }
        }

        internal static void Shutdown()
        {
            UnregisterKeyboardHotkeys();
            DisposeExternalEvents();

            if (_hotkeySource != null)
            {
                _hotkeySource.RemoveHook(HotkeyWindowHook);
                _hotkeySource.Dispose();
                _hotkeySource = null;
            }

            _isInitialized = false;
            _isSuspended = false;
            _lastRaisedTarget = null;
        }

        internal static void SuspendHotkeys()
        {
            _isSuspended = true;
            UnregisterKeyboardHotkeys();
        }

        internal static void ResumeHotkeys()
        {
            _isSuspended = false;

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

        private static void DisposeExternalEvents()
        {
            ExternalEvent commandSearchEvent = _openCommandSearchEvent;
            ExternalEvent commandsWheelEvent = _openCommandsWheelEvent;
            _openCommandSearchEvent = null;
            _openCommandsWheelEvent = null;

            TryDisposeExternalEvent(commandSearchEvent);
            TryDisposeExternalEvent(commandsWheelEvent);
        }

        private static void TryDisposeExternalEvent(ExternalEvent externalEvent)
        {
            if (externalEvent == null)
            {
                return;
            }

            try
            {
                externalEvent.Dispose();
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

            if (HotkeyGestureService.IsEmpty(gesture))
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

            Key parsedKey;
            if (!TryParseKey(mainKeys[0], out parsedKey))
            {
                return false;
            }

            if (modifiers == 0 && parsedKey != Key.Tab)
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
            if (_isSuspended)
            {
                return IntPtr.Zero;
            }

            if (message != WM_HOTKEY)
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

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int processId);
    }

}