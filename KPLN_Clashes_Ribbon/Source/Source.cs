using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using static KPLN_Clashes_Ribbon.Common.Collections;

namespace KPLN_Clashes_Ribbon.Source
{
    public class Source
    {
        private static string AssemblyPath = new FileInfo(Assembly.GetExecutingAssembly().Location).DirectoryName;
        
        public Source(Icon icon)
        {
            switch (icon)
            {
                case Icon.Report:
                    Value = Path.Combine(AssemblyPath, @"Source\ImageData\report_default.png");
                    break;
                case Icon.Report_Closed:
                    Value = Path.Combine(AssemblyPath, @"Source\ImageData\report_closed.png");
                    break;
                case Icon.Report_New:
                    Value = Path.Combine(AssemblyPath, @"Source\ImageData\report_new.png");
                    break;
                case Icon.Instance:
                    Value = Path.Combine(AssemblyPath, @"Source\ImageData\report_instance.png");
                    break;
                case Icon.Instance_Closed:
                    Value = Path.Combine(AssemblyPath, @"Source\ImageData\report_instance_closed.png");
                    break;
                case Icon.Instance_Delegated:
                    Value = Path.Combine(AssemblyPath, @"Source\ImageData\report_instance_delegated.png");
                    break;
            }
        }

        public string Value { get; }
    }
}
