using System;

namespace KPLN_UserDataAgent.Services
{
    internal sealed class UserEventRecord
    {
        public long LocalId { get; set; }
        public string SyncId { get; set; }
        public string EventTime { get; set; }
        public string WindowsUser { get; set; }
        public int SubDepartmentId { get; set; }
        public int RevitVersion { get; set; }
        public string DocumentTitle { get; set; }
        public string DocumentPath { get; set; }
        public string EventName { get; set; }
        public string TransactionName { get; set; }
        public int AddedCount { get; set; }
        public int ModifiedCount { get; set; }
        public int DeletedCount { get; set; }

        public static UserEventRecord Create(
            string eventName,
            string transactionName,
            DocumentSnapshot document,
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
                SubDepartmentId = userContext.SubDepartmentId,
                RevitVersion = ModuleData.RevitVersion,
                DocumentTitle = document == null ? string.Empty : document.Title ?? string.Empty,
                DocumentPath = document == null ? string.Empty : document.Path ?? string.Empty,
                EventName = eventName ?? string.Empty,
                TransactionName = transactionName ?? string.Empty,
                AddedCount = addedCount,
                ModifiedCount = modifiedCount,
                DeletedCount = deletedCount
            };
        }
    }
}
