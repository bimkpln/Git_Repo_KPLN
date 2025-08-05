namespace KPLN_Clashes_Ribbon.Core
{
    public static class ClashesMainCollection
    {
        public enum KPIcon { Default, Report, Report_New, Report_Closed, Instance, Instance_Closed, Instance_Delegated }

        public enum KPItemStatus { New, Opened, Closed, Approved, Delegated }

        public enum KPTaskDialogResult { Ok, Cancel, None }

        public enum KPTaskDialogIcon { Sad, Happy, Warning, Question, Lol, Ooo }

        public static readonly string StringSeparatorItem = "~SE00~";

        public static readonly string StringSeparatorSubItem = "~SE01~";
    }
}
