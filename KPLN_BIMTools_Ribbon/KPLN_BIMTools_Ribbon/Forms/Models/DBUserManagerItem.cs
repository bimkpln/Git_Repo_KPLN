using KPLN_Library_DBWorker.Core;

namespace KPLN_BIMTools_Ribbon.Forms.Models
{
    /// <summary>
    /// Представление пользователя БД вместе с читаемым именем подразделения.
    /// </summary>
    public sealed class DBUserManagerItem
    {
        public DBUserManagerItem(DBUser dbUser, string subDepartmentName)
        {
            DBUser = dbUser;
            SubDepartmentName = string.IsNullOrWhiteSpace(subDepartmentName)
                ? $"ID {dbUser.SubDepartmentId}"
                : subDepartmentName;
        }

        public DBUser DBUser { get; }

        public int Id => DBUser.Id;

        public string SystemName => DBUser.SystemName;

        public string FullName => $"{DBUser.Surname} {DBUser.Name}".Trim();

        public int SubDepartmentId => DBUser.SubDepartmentId;

        public string SubDepartmentName { get; }

        public string RegistrationDate => DBUser.RegistrationDate;

        public string LastConnectionDate => DBUser.LastConnectionDate;

        public string RevitUserName => DBUser.RevitUserName;

        public bool IsDebugMode => DBUser.IsDebugMode;

        public bool IsUserRestricted => DBUser.IsUserRestricted;

        public int BitrixUserID => DBUser.BitrixUserID;

        public bool IsExtraNet => DBUser.IsExtraNet;

        public bool IsFired => DBUser.IsFired;

        public string AccessStatus => IsUserRestricted ? "Закрыт" : "Открыт";
    }
}
