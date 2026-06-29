using System.Collections.Generic;
using System.Runtime.Serialization;

namespace KPLN_CommandsWheel.Models
{
    [DataContract]
    public class UserSettings
    {
        [DataMember]
        public List<string> FavoriteCommandIds { get; set; } = new List<string>();

        [DataMember]
        public List<string> WheelCommandIds { get; set; } = new List<string>();

        [DataMember]
        public List<string> RecentCommandIds { get; set; } = new List<string>();
    }
}
