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
        private const int WM_INPUT = 0x00FF;
        private const int WM_HOTKEY = 0x0312;
        private const int RIM_TYPEMOUSE = 0;
        private const int CommandSearchHotkeyId = 0x4B50;
        private const int CommandsWheelHotkeyId = 0x4B51;
        private const ushort HID_USAGE_PAGE_GENERIC = 0x01;
        private const ushort HID_USAGE_GENERIC_MOUSE = 0x02;
        private const ushort RI_MOUSE_BUTTON_4_DOWN = 0x0040;
        private const ushort RI_MOUSE_BUTTON_5_DOWN = 0x0100;
        private const uint MOD_ALT = 0x0001;
        private const uint MOD_CONTROL = 0x0002;
        private const uint MOD_SHIFT = 0x0004;
        private const uint MOD_NOREPEAT = 0x4000;
        private const uint RID_INPUT = 0x10000003;
        private const uint RIDEV_REMOVE = 0x00000001;
        private const uint RIDEV_INPUTSINK = 0x00000100;

        private static ExternalEvent _openCommandSearchEvent;
        private static ExternalEvent _openCommandsWheelEvent;
        private static HwndSource _hotkeySource;
        private static UserSettings _settings;
        private static bool _isInitialized;
        private static bool _isSuspended;
        private static bool _isRawMouseRegistered;
        private static IntPtr _rawInputBuffer = IntPtr.Zero;
        private static uint _rawInputBufferSize;
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
            UpdateInputHandlers();
        }

        internal static void ReloadSettings(UserSettings settings = null)
        {
            _settings = settings ?? UserSettingsService.Load();

            if (_isInitialized)
            {
                RegisterKeyboardHotkeys();
                UpdateInputHandlers();
            }
        }

        internal static void Shutdown()
        {
            UnregisterKeyboardHotkeys();
            UnregisterRawMouseInput();
            FreeRawInputBuffer();

            if (_hotkeySource != null)
            {
                _hotkeySource.RemoveHook(HotkeyWindowHook);
                _hotkeySource.Dispose();
                _hotkeySource = null;
            }

            _isInitialized = false;
            _isSuspended = false;
        }

        internal static void SuspendHotkeys()
        {
            _isSuspended = true;
            UnregisterKeyboardHotkeys();
            UnregisterRawMouseInput();
        }

        internal static void ResumeHotkeys()
        {
            _isSuspended = false;

            if (_isInitialized)
            {
                RegisterKeyboardHotkeys();
                UpdateInputHandlers();
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

        private static void UpdateInputHandlers()
        {
            if (!_isInitialized)
            {
                return;
            }

            if (_isSuspended || _settings == null)
            {
                UnregisterRawMouseInput();
                return;
            }

            if (NeedsRawMouseInput())
            {
                RegisterRawMouseInput();
            }
            else
            {
                UnregisterRawMouseInput();
            }
        }

        private static bool NeedsRawMouseInput()
        {
            return HasMouseButton(_settings.CommandSearchHotkey)
                || HasMouseButton(_settings.CommandsWheelHotkey);
        }

        private static bool HasMouseButton(HotkeyGesture gesture)
        {
            return gesture != null && !string.IsNullOrWhiteSpace(gesture.MouseButton);
        }

        private static void RegisterRawMouseInput()
        {
            if (_isRawMouseRegistered)
            {
                return;
            }

            EnsureHotkeyWindow();
            if (_hotkeySource == null || _hotkeySource.Handle == IntPtr.Zero)
            {
                return;
            }

            RawInputDevice[] devices =
            {
                new RawInputDevice
                {
                    UsagePage = HID_USAGE_PAGE_GENERIC,
                    Usage = HID_USAGE_GENERIC_MOUSE,
                    Flags = RIDEV_INPUTSINK,
                    Target = _hotkeySource.Handle
                }
            };

            _isRawMouseRegistered = RegisterRawInputDevices(
                devices,
                (uint)devices.Length,
                (uint)Marshal.SizeOf(typeof(RawInputDevice)));
        }

        private static void UnregisterRawMouseInput()
        {
            if (!_isRawMouseRegistered)
            {
                return;
            }

            RawInputDevice[] devices =
            {
                new RawInputDevice
                {
                    UsagePage = HID_USAGE_PAGE_GENERIC,
                    Usage = HID_USAGE_GENERIC_MOUSE,
                    Flags = RIDEV_REMOVE,
                    Target = IntPtr.Zero
                }
            };

            RegisterRawInputDevices(
                devices,
                (uint)devices.Length,
                (uint)Marshal.SizeOf(typeof(RawInputDevice)));

            _isRawMouseRegistered = false;
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

            if (message == WM_INPUT)
            {
                if (_isRawMouseRegistered && IsForegroundCurrentProcess() && TryHandleRawMouseInput(lParam))
                {
                    handled = true;
                }

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

        private static bool TryHandleRawMouseInput(IntPtr lParam)
        {
            uint size = 0;
            uint headerSize = (uint)Marshal.SizeOf(typeof(RawInputHeader));

            if (GetRawInputData(lParam, RID_INPUT, IntPtr.Zero, ref size, headerSize) == uint.MaxValue || size == 0)
            {
                return false;
            }

            if (!EnsureRawInputBuffer(size))
            {
                return false;
            }

            uint readSize = GetRawInputData(lParam, RID_INPUT, _rawInputBuffer, ref size, headerSize);
            if (readSize == uint.MaxValue || readSize != size)
            {
                return false;
            }

            RawInput rawInput = (RawInput)Marshal.PtrToStructure(_rawInputBuffer, typeof(RawInput));
            if (rawInput.Header.Type != RIM_TYPEMOUSE)
            {
                return false;
            }

            string mouseButton = null;
            if ((rawInput.Mouse.ButtonFlags & RI_MOUSE_BUTTON_4_DOWN) == RI_MOUSE_BUTTON_4_DOWN)
            {
                mouseButton = "XButton1";
            }
            else if ((rawInput.Mouse.ButtonFlags & RI_MOUSE_BUTTON_5_DOWN) == RI_MOUSE_BUTTON_5_DOWN)
            {
                mouseButton = "XButton2";
            }

            return !string.IsNullOrWhiteSpace(mouseButton) && TryTriggerMouseHotkey(mouseButton);
        }

        private static bool EnsureRawInputBuffer(uint size)
        {
            if (size == 0 || size > int.MaxValue)
            {
                return false;
            }

            if (_rawInputBuffer != IntPtr.Zero && _rawInputBufferSize >= size)
            {
                return true;
            }

            FreeRawInputBuffer();

            _rawInputBuffer = Marshal.AllocHGlobal((int)size);
            _rawInputBufferSize = size;
            return _rawInputBuffer != IntPtr.Zero;
        }

        private static void FreeRawInputBuffer()
        {
            if (_rawInputBuffer == IntPtr.Zero)
            {
                return;
            }

            Marshal.FreeHGlobal(_rawInputBuffer);
            _rawInputBuffer = IntPtr.Zero;
            _rawInputBufferSize = 0;
        }

        private static bool TryTriggerMouseHotkey(string mouseButton)
        {
            if (_settings == null)
            {
                ReloadSettings();
            }

            List<string> pressedKeys = GetPressedModifierKeys();

            if (HotkeyGestureService.Matches(_settings.CommandSearchHotkey, pressedKeys, mouseButton))
            {
                Raise(HotkeyTarget.CommandSearch);
                return true;
            }

            if (HotkeyGestureService.Matches(_settings.CommandsWheelHotkey, pressedKeys, mouseButton))
            {
                Raise(HotkeyTarget.CommandsWheel);
                return true;
            }

            return false;
        }

        private static List<string> GetPressedModifierKeys()
        {
            List<string> keys = new List<string>();
            ModifierKeys modifiers = Keyboard.Modifiers;

            if ((modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                keys.Add("Ctrl");
            }

            if ((modifiers & ModifierKeys.Shift) == ModifierKeys.Shift)
            {
                keys.Add("Shift");
            }

            if ((modifiers & ModifierKeys.Alt) == ModifierKeys.Alt)
            {
                keys.Add("Alt");
            }

            return HotkeyGestureService.NormalizeKeys(keys);
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

        [StructLayout(LayoutKind.Sequential)]
        private struct RawInputDevice
        {
            public ushort UsagePage;
            public ushort Usage;
            public uint Flags;
            public IntPtr Target;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RawInputHeader
        {
            public uint Type;
            public uint Size;
            public IntPtr Device;
            public IntPtr WParam;
        }

        [StructLayout(LayoutKind.Explicit)]
        private struct RawMouse
        {
            [FieldOffset(0)]
            public ushort Flags;

            [FieldOffset(4)]
            public ushort ButtonFlags;

            [FieldOffset(6)]
            public ushort ButtonData;

            [FieldOffset(8)]
            public uint RawButtons;

            [FieldOffset(12)]
            public int LastX;

            [FieldOffset(16)]
            public int LastY;

            [FieldOffset(20)]
            public uint ExtraInformation;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RawInput
        {
            public RawInputHeader Header;
            public RawMouse Mouse;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RegisterRawInputDevices(
            RawInputDevice[] pRawInputDevices,
            uint uiNumDevices,
            uint cbSize);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetRawInputData(
            IntPtr hRawInput,
            uint uiCommand,
            IntPtr pData,
            ref uint pcbSize,
            uint cbSizeHeader);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int processId);
    }
}