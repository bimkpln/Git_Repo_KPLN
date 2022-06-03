using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using static KPLN_ModelChecker_User.Common.Collections;

namespace KPLN_ModelChecker_User.Source
{
    public class Source
    {
        public string Value { get; }
        private static string AssemblyPath = new FileInfo(Assembly.GetExecutingAssembly().Location).DirectoryName;
        public Source(Icon icon)
        {
            switch (icon)
            {
                case Icon.CheckLevels:
                    Value = Path.Combine(AssemblyPath, @"Source\checker_levels.png");
                    break;
                case Icon.PullDown:
                    Value = Path.Combine(AssemblyPath, @"Source\checker_pull.png");
                    break;
                case Icon.CheckMirrored:
                    Value = Path.Combine(AssemblyPath, @"Source\checker_mirrored.png");
                    break;
                case Icon.CheckLocations:
                    Value = Path.Combine(AssemblyPath, @"Source\checker_locations.png");
                    break;
                case Icon.CheckMonitorGrids:
                    Value = Path.Combine(AssemblyPath, @"Source\checker_grids_monitor.png");
                    break;
                case Icon.CheckMonitorLevels:
                    Value = Path.Combine(AssemblyPath, @"Source\checker_levels_monitor.png");
                    break;
                case Icon.Errors:
                    Value = Path.Combine(AssemblyPath, @"Source\checker_push.png");
                    break;
                case Icon.FamilyName:
                    Value = Path.Combine(AssemblyPath, @"Source\family_name.png");
                    break;
                case Icon.Worksets:
                    Value = Path.Combine(AssemblyPath, @"Source\checker_worksets.png");
                    break;
            }
        }
    }
}
