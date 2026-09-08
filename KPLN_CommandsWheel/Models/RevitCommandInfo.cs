using System.Runtime.Serialization;
using System.Windows.Media;

namespace KPLN_CommandsWheel.Models
{
    [DataContract]
    public class RevitCommandInfo
    {
        [DataMember(Order = 1)]
        public string Id { get; set; }

        [DataMember(Order = 2)]
        public string Name { get; set; }

        [DataMember(Order = 3)]
        public string TabName { get; set; }

        [DataMember(Order = 4)]
        public string PanelName { get; set; }

        [DataMember(Order = 5)]
        public string Tooltip { get; set; }

        // Store image bytes, not a session-only WPF object or a resource URI.
        [DataMember(Order = 6, EmitDefaultValue = false)]
        public string IconPngBase64 { get; set; }

        public ImageSource RibbonImage { get; set; }

        public string SearchText
        {
            get
            {
                return string.Format(
                    "{0} {1} {2} {3} {4}",
                    Id,
                    Name,
                    TabName,
                    PanelName,
                    Tooltip
                );
            }
        }
    }
}