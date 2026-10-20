using System.Reflection;

namespace KPLN_ViewsAndLists_CopyByLevels
{
    internal static class ModuleData
    {
        public static readonly string ModuleName = Assembly.GetExecutingAssembly().GetName().Name;
        public static readonly string Version = Assembly.GetExecutingAssembly().GetName().Version.ToString();
        public static int RevitVersion { get; set; }
    }
}
