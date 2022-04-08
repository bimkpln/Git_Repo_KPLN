using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_NavisWorksReports.Common
{
    public static class Collections
    {
        public enum Icon { Default, Report, Report_New, Report_Closed, Instance, Instance_Closed }
        public enum Status { Opened, Closed, Approved }
        public static readonly string separator_element = "~SE00~";
        public static readonly string separator_sub_element = "~SE01~";
        public enum KPTaskDialogResult { Ok, Cancel, None }
        public enum KPTaskDialogIcon { Sad, Happy, Warning, Question, Lol, Ooo }
    }
}
