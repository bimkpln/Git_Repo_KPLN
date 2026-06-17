using KPLN_Library_DBWorker;
using KPLN_Library_DBWorker.Core;
using KPLN_Loader.Core.Entities;

namespace KPLN_UserDataAgent.Services
{
    internal sealed class UserContextSnapshot
    {
        public string UserName { get; private set; }
        public string DepartmentKey { get; private set; }

        public static UserContextSnapshot Current()
        {
            string userName = "DB_USER_NOT_FOUND";

            try
            {
                DBUser user = SQLiteMainService.CurrentDBUser;
                if (user != null)
                {
                    if (!string.IsNullOrWhiteSpace(user.SystemName))
                        userName = user.SystemName;
                    else
                    {
                        string fullName = string.Format("{0} {1}", user.Name, user.Surname).Trim();
                        if (!string.IsNullOrWhiteSpace(fullName))
                            userName = fullName;
                    }
                }
            }
            catch
            {
            }

            return new UserContextSnapshot
            {
                UserName = userName,
                DepartmentKey = ReferenceDepartmentLookupService.ResolveDepartmentKey(
                    userName,
                    ModuleData.ReferenceDatabasePath)
            };
        }
    }
}
