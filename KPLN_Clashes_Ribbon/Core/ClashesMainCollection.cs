using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;

namespace KPLN_Clashes_Ribbon.Core
{
    public static class ClashesMainCollection
    {
        private static UserDbService _clashUserDbService;

        private static DBUser _currentDBUser;

        public enum KPIcon { Default, Report, Report_New, Report_Closed, Instance, Instance_Closed, Instance_Delegated }

        public enum KPItemStatus { New, Opened, Closed, Approved, Delegated }

        public enum KPTaskDialogResult { Ok, Cancel, None }

        public enum KPTaskDialogIcon { Sad, Happy, Warning, Question, Lol, Ooo }

        public static readonly string StringSeparatorItem = "~SE00~";

        public static readonly string StringSeparatorSubItem = "~SE01~";

        public static UserDbService ClashUserDbService
        {
            get
            {
                if (_clashUserDbService == null) 
                    _clashUserDbService = (UserDbService)new CreatorUserDbService().CreateService();

                return _clashUserDbService;
            }
        }

        public static DBUser CurrentDBUser 
        {
            get
            {
                if (_currentDBUser == null)
                    _currentDBUser = ClashUserDbService.GetCurrentDBUser();
                
                return _currentDBUser;
            } 
        }

        public static DBUser GetDBUser_ByUserName(string sysUserName) => ClashUserDbService.GetDBUser_ByUserName(sysUserName);
    }
}
