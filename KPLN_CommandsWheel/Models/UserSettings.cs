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
        public HotkeyGesture CommandSearchHotkey { get; set; } = new HotkeyGesture();

        [DataMember]
        public HotkeyGesture CommandsWheelHotkey { get; set; } = new HotkeyGesture();
    }

    [DataContract]
    public class HotkeyGesture
    {
        [DataMember]
        public List<string> Keys { get; set; } = new List<string>();

        [DataMember]
        public string MouseButton { get; set; }
    }
}