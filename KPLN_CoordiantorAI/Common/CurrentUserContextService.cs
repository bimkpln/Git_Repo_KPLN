using KPLN_Library_DBWorker;
using System;

namespace KPLN_CoordiantorAI.Common
{
    public sealed class CurrentUserContextService
    {
        public CurrentUserContext GetCurrentUserContext()
        {
            CurrentUserContext context = new CurrentUserContext
            {
                UserName = Environment.UserName,
                SubDepartmentId = -1
            };

            try
            {
                var currentUser = SQLiteMainService.CurrentDBUser;
                if (currentUser != null)
                {
                    string fullName = string.Format("{0} {1}", currentUser.Name, currentUser.Surname).Trim();
                    if (!string.IsNullOrWhiteSpace(fullName))
                        context.UserName = fullName;
                    else if (!string.IsNullOrWhiteSpace(currentUser.SystemName))
                        context.UserName = currentUser.SystemName;

                    if (currentUser.SubDepartmentId > 0)
                        context.SubDepartmentId = currentUser.SubDepartmentId;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                var subDepartment = SQLiteMainService.CurrentUserDBSubDepartment;
                if (subDepartment != null && subDepartment.Id > 0)
                    context.SubDepartmentId = subDepartment.Id;
            }
            catch (Exception)
            {
            }

            return context;
        }
    }
}