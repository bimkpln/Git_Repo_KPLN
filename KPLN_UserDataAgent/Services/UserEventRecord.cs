using System;

namespace KPLN_UserDataAgent.Services
{
    internal sealed class UserEventRecord
    {
        public long LocalId { get; set; }
        public string LocalDatabasePath { get; set; }
        public string SyncId { get; set; }
        public string EventTime { get; set; }
        public string WindowsUser { get; set; }
        public string DepartmentKey { get; set; }
        public long EventTransactionId { get; set; }
        public string EventName { get; set; }
        public string TransactionName { get; set; }
        public int AddedCount { get; set; }
        public int ModifiedCount { get; set; }
        public int DeletedCount { get; set; }

        public static UserEventRecord Create(
            string eventName,
            string transactionName,
            UserContextSnapshot userContext,
            int addedCount = 0,
            int modifiedCount = 0,
            int deletedCount = 0)
        {
            return new UserEventRecord
            {
                SyncId = Guid.NewGuid().ToString("N"),
                EventTime = DateTime.Now.ToString("yyyy.MM.dd. HH:mm:ss"),
                WindowsUser = userContext.UserName ?? string.Empty,
                DepartmentKey = userContext.DepartmentKey ?? CentralDatabasePathBuilder.UnknownDepartmentKey,
                EventName = eventName ?? string.Empty,
                TransactionName = transactionName ?? string.Empty,
                AddedCount = addedCount,
                ModifiedCount = modifiedCount,
                DeletedCount = deletedCount
            };
        }
    }
}