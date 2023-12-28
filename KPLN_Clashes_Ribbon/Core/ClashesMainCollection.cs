using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;

namespace KPLN_Clashes_Ribbon.Core
{
    public static class ClashesMainCollection
    {
        private static DBUser _dBUser;

        public static readonly string StringSeparatorItem = "~SE00~";

        public static readonly string StringSeparatorSubItem = "~SE01~";

        public static DBUser CurrentDBUser 
        {
            get
            {
                if (_dBUser == null)
                {
                    UserDbService userDbService = (UserDbService) new CreatorUserDbService().CreateService();
                    _dBUser = userDbService.GetCurrentDBUser();
                }
                return _dBUser;
            } 
        }

        public enum KPIcon { Default, Report, Report_New, Report_Closed, Instance, Instance_Closed, Instance_Delegated }

        public enum KPItemStatus { New, Opened, Closed, Approved, Delegated }

        public enum KPTaskDialogResult { Ok, Cancel, None }

        public enum KPTaskDialogIcon { Sad, Happy, Warning, Question, Lol, Ooo }

    }
}
