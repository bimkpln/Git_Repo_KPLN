using KPLN_Loader.Common;
using KPLN_DataBase.Collections;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace KPLN_NavisWorksReports
{
    public static class ModuleData
    {
#if Revit2022
        public static string RevitVersion = "2022";
        public static Window RevitWindow { get; set; }
#endif
#if Revit2020
        public static string RevitVersion = "2020";
        public static Window RevitWindow { get; set; }
#endif
#if Revit2018
        public static string RevitVersion = "2018";
#endif
        public static System.IntPtr MainWindowHandle { get; set; }
        public static string Build = string.Format("built for Revit {0}", RevitVersion);
        public static string Version = "1.0.2.3b";
        public static string Date = "2020/12/22";
        public static string ModuleName = "NavisWorks Reports";
        public static List<SQLUserInfo> Users = KPLN_Loader.Preferences.Users;
        public static string GetUserBySystemName(string name)
        {
            foreach (SQLUserInfo user in Users)
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
