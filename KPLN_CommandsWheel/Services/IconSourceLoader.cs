using KPLN_CommandsWheel.Models;
using System.Windows.Media;

namespace KPLN_CommandsWheel.Services
{
    internal static class IconSourceLoader
    {
        internal static ImageSource Load(RevitCommandInfo command)
        {
            return command == null ? null : command.RibbonImage;
        }
    }
}