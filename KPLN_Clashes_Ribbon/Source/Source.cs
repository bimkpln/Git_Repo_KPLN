using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using static KPLN_Clashes_Ribbon.Core.ClashesMainCollection;

namespace KPLN_Clashes_Ribbon.Source
{
    public class Source
    {
        private static readonly string AssemblyPath = new FileInfo(Assembly.GetExecutingAssembly().Location).DirectoryName;
        
        public Source(KPIcon icon)
        {
            switch (icon)
            {
                case KPIcon.Report:
                    Value = Path.Combine(AssemblyPath, @"Source\ImageData\report_default.png");
                    break;
                case KPIcon.Report_Closed:
                    Value = Path.Combine(AssemblyPath, @"Source\ImageData\report_closed.png");
                    break;
                case KPIcon.Report_New:
                    Value = Path.Combine(AssemblyPath, @"Source\ImageData\report_new.png");
                    break;
                case KPIcon.Instance:
                    Value = Path.Combine(AssemblyPath, @"Source\ImageData\report_instance.png");
                    break;
                case KPIcon.Instance_Closed:
                    Value = Path.Combine(AssemblyPath, @"Source\ImageData\report_instance_closed.png");
                    break;
                case KPIcon.Instance_Delegated:
                    Value = Path.Combine(AssemblyPath, @"Source\ImageData\report_instance_delegated.png");
                    break;
                case KPIcon.Instance_Approved:
                    Value = Path.Combine(AssemblyPath, @"Source\ImageData\report_instance_approved.png");
                    break;
            }
        }

        public string Value { get; }
    }
}
