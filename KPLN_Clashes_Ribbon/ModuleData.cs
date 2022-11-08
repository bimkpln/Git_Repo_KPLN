using KPLN_Loader.Common;
using System.Reflection;

namespace KPLN_Clashes_Ribbon
{
    internal static class ModuleData
    {
        /// <summary>
        /// Версия сборки, отображаемая в Revit
        /// </summary>
        public static string Version = Assembly.GetExecutingAssembly().GetName().Version.ToString();

        /// <summary>
        /// Актуальная дата плагина
        /// </summary>
        public static string Date = "2022/11/07";

        /// <summary>
        /// Имя модуля
        /// </summary>
        public static string ModuleName = Assembly.GetExecutingAssembly().GetName().Name;

#if Revit2020
        public static string RevitVersion = "2020";
#endif
#if Revit2018
        public static string RevitVersion = "2018";
#endif

        public static string GetUserBySystemName(string name)
        {
            foreach (SQLUserInfo user in KPLN_Loader.Preferences.Users)
            {
                if (user.SystemName == name)
                {
                    if (user.Surname != string.Empty)
                    {
                        return string.Format("{0} {1}.{2}.", user.Family, user.Name[0], user.Surname[0]);
                    }
                    else
                    {
                        return string.Format("{0} {1}", user.Family, user.Name);
                    }
                }
            }
            return name;
        }
    }

}
