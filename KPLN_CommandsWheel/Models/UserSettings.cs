using System.Collections.Generic;
using System.Runtime.Serialization;

namespace KPLN_CommandsWheel.Models
{
    public static class WheelModeNames
    {
        public const string Unpinned = "Unpinned";
        public const string Pinned = "Pinned";
    }

    [DataContract]
    public class UserSettings
    {
        [DataMember]
        public List<string> FavoriteCommandIds { get; set; } = new List<string>();

        [DataMember]
        public List<string> WheelCommandIds { get; set; } = new List<string>();

        [DataMember]
        public List<string> RecentCommandIds { get; set; } = new List<string>();

        [DataMember]
        public string WheelMode { get; set; } = WheelModeNames.Unpinned;

        [DataMember]
        public bool IsWheelCloseButtonVisible { get; set; }

        [DataMember]
        public string WheelShortcut { get; set; } = string.Empty;

        [DataMember]
        public string CommandSearchShortcut { get; set; } = string.Empty;

        [DataMember]
        public bool AreKeyboardShortcutsConfigured { get; set; }

        [DataMember]
        public LegacyHotkeyGesture CommandSearchHotkey { get; set; }

        [DataMember]
        public LegacyHotkeyGesture CommandsWheelHotkey { get; set; }

        [DataMember]
        public bool LegacyHotkeyMigrationAttempted { get; set; }

        [DataMember]
        public bool LegacyHotkeyMigrationNoticeShown { get; set; }

        [DataMember]
        public string KeyboardShortcutMigrationStatus { get; set; } = "Pending";

        [DataMember]
        public string KeyboardShortcutMigrationMessage { get; set; }
    }

    [DataContract]
    public class LegacyHotkeyGesture
    {
        [DataMember]
        public List<string> Keys { get; set; } = new List<string>();

        [DataMember]
        public string MouseButton { get; set; }
    }
}