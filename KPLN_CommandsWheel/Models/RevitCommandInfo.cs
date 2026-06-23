using System.Windows.Media;

namespace KPLN_CommandsWheel.Models
{
    public class RevitCommandInfo
    {
        public string Id { get; set; }

        public string Name { get; set; }

        public string TabName { get; set; }

        public string PanelName { get; set; }

        public string Tooltip { get; set; }

        public bool CanPost { get; set; }

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