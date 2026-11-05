using KPLN_Library_DBWorker;

namespace KPLN_ModelChecker_User.Checks.KR
{
    internal static class KRCheckAccess
    {
        internal static bool IsAllowed()
        {
            // Unknown user / unavailable DB must not break the entire ribbon module.
            try
            {
                var user = SQLiteMainService.CurrentDBUser;
                return user != null && IsAllowedDepartment(user.SubDepartmentId);
            }
            catch { return false; }
        }

        internal static bool IsAllowedDepartment(int departmentId)
        {
            // SubDepartments: 3 = KR, 8 = BIM. Debug status does not grant access.
            return departmentId == 3 || departmentId == 8;
        }
    }
}
