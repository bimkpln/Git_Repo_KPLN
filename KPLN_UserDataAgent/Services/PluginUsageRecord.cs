using System;

namespace KPLN_UserDataAgent.Services
{
    internal sealed class PluginUsageRecord
    {
        public long LocalId { get; set; }
        public string LocalDatabasePath { get; set; }
        public string SyncId { get; set; }
        public string RunId { get; set; }
        public string EventType { get; set; }
        public string EventTime { get; set; }
        public string WindowsUser { get; set; }
        public string DepartmentKey { get; set; }
        public string TabName { get; set; }
        public string PanelName { get; set; }
        public string ButtonName { get; set; }
        public string TransactionName { get; set; }
        public int AddedCount { get; set; }
        public int ModifiedCount { get; set; }
        public int DeletedCount { get; set; }

        public static PluginUsageRecord Create(
            string runId,
            string eventType,
            string tabName,
            string panelName,
            string buttonName,
            string transactionName,
            UserContextSnapshot userContext,
            int addedCount = 0,
            int modifiedCount = 0,
            int deletedCount = 0)
        {
            return new PluginUsageRecord
            {
                SyncId = Guid.NewGuid().ToString("N"),
                RunId = string.IsNullOrWhiteSpace(runId) ? Guid.NewGuid().ToString("N") : runId,
                EventType = eventType ?? string.Empty,
                EventTime = DateTime.Now.ToString("yyyy.MM.dd. HH:mm:ss"),
                WindowsUser = userContext.UserName ?? string.Empty,
                DepartmentKey = userContext.DepartmentKey ?? CentralDatabasePathBuilder.UnknownDepartmentKey,
                TabName = tabName ?? string.Empty,
                PanelName = panelName ?? string.Empty,
                ButtonName = buttonName ?? string.Empty,
                TransactionName = transactionName ?? string.Empty,
                AddedCount = addedCount,
                ModifiedCount = modifiedCount,
                DeletedCount = deletedCount
            };
        }
    }
}