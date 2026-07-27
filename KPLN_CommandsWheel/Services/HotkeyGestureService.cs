using KPLN_CommandsWheel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Input;

namespace KPLN_CommandsWheel.Services
{
    internal static class HotkeyGestureService
    {
        private static readonly string[] ModifierOrder = { "Ctrl", "Shift", "Alt" };

        internal static bool IsEmpty(HotkeyGesture gesture)
        {
            return gesture == null
                || gesture.Keys == null
                || gesture.Keys.Count == 0;
        }

        internal static string ToDisplayText(HotkeyGesture gesture)
        {
            if (IsEmpty(gesture))
            {
                return "Не назначено";
            }

            return string.Join(" + ", NormalizeKeys(gesture.Keys));
        }

        internal static bool AreEqual(HotkeyGesture first, HotkeyGesture second)
        {
            if (IsEmpty(first) && IsEmpty(second))
            {
                return true;
            }

            if (IsEmpty(first) || IsEmpty(second))
            {
                return false;
            }

            return NormalizeKeys(first.Keys).SequenceEqual(NormalizeKeys(second.Keys), StringComparer.OrdinalIgnoreCase);
        }

        internal static HotkeyGesture FromKeyboardEvent(KeyEventArgs args)
        {
            if (args == null)
            {
                return new HotkeyGesture();
            }

            List<string> keys = GetModifierKeys();
            string keyName = NormalizeKey(GetActualKey(args));

            if (!string.IsNullOrWhiteSpace(keyName) && !IsModifier(keyName))
            {
                keys.Add(keyName);
            }

            return new HotkeyGesture { Keys = NormalizeKeys(keys) };
        }

        internal static string GetKeyName(KeyEventArgs args)
        {
            return args == null ? null : NormalizeKey(GetActualKey(args));
        }

        internal static string NormalizeKey(Key key)
        {
            switch (key)
            {
                case Key.LeftCtrl:
                case Key.RightCtrl:
                    return "Ctrl";

                case Key.LeftShift:
                case Key.RightShift:
                    return "Shift";

                case Key.LeftAlt:
                case Key.RightAlt:
                    return "Alt";

                case Key.Return:
                    return "Enter";

                case Key.None:
                    return null;

                default:
                    return key.ToString();
            }
        }

        internal static List<string> NormalizeKeys(IEnumerable<string> keys)
        {
            if (keys == null)
            {
                return new List<string>();
            }

            return keys
                .Select(NormalizeKeyName)
                .Where(key => !string.IsNullOrWhiteSpace(key))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(GetKeySortOrder)
                .ThenBy(key => key, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<string> GetModifierKeys()
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

            return keys;
        }

        private static Key GetActualKey(KeyEventArgs args)
        {
            if (args.Key == Key.System)
            {
                return args.SystemKey;
            }

            if (args.Key == Key.ImeProcessed)
            {
                return args.ImeProcessedKey;
            }

            return args.Key;
        }

        private static string NormalizeKeyName(string key)
        {
            if (string.IsNullOrWhiteSpace(key))
            {
                return null;
            }

            string value = key.Trim();
            if (string.Equals(value, "Control", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "LeftCtrl", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "RightCtrl", StringComparison.OrdinalIgnoreCase))
            {
                return "Ctrl";
            }

            if (string.Equals(value, "LeftShift", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "RightShift", StringComparison.OrdinalIgnoreCase))
            {
                return "Shift";
            }

            if (string.Equals(value, "Menu", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "LeftAlt", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "RightAlt", StringComparison.OrdinalIgnoreCase))
            {
                return "Alt";
            }

            if (string.Equals(value, "Return", StringComparison.OrdinalIgnoreCase))
            {
                return "Enter";
            }

            return value;
        }

        internal static bool IsModifier(string key)
        {
            return ModifierOrder.Any(modifier => string.Equals(modifier, key, StringComparison.OrdinalIgnoreCase));
        }

        private static int GetKeySortOrder(string key)
        {
            for (int index = 0; index < ModifierOrder.Length; index++)
            {
                if (string.Equals(ModifierOrder[index], key, StringComparison.OrdinalIgnoreCase))
                {
                    return index;
                }
            }

            return ModifierOrder.Length;
        }
    }
}
