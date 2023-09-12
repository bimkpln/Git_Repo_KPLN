namespace KPLN_Clashes_Ribbon.Core
{
    public static class ClashesMainCollection
    {
        public static readonly string separator_element = "~SE00~";

        public static readonly string separator_sub_element = "~SE01~";

        public enum KPIcon { Default, Report, Report_New, Report_Closed, Instance, Instance_Closed, Instance_Delegated }

        public enum KPItemStatus { New, Opened, Closed, Approved, Delegated }

        public enum KPTaskDialogResult { Ok, Cancel, None }

        public enum KPTaskDialogIcon { Sad, Happy, Warning, Question, Lol, Ooo }
    }
}
