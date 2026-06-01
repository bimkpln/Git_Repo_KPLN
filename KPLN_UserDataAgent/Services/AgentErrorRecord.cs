using System;

namespace KPLN_UserDataAgent.Services
{
    internal sealed class AgentErrorRecord
    {
        private const int MaxTextLength = 8000;

        public long LocalId { get; set; }
        public string LocalDatabasePath { get; set; }
        public string SyncId { get; set; }
        public string ErrorTime { get; set; }
        public string WindowsUser { get; set; }
        public int SubDepartmentId { get; set; }
        public int RevitVersion { get; set; }
        public string Source { get; set; }
        public string ErrorType { get; set; }
        public string ErrorMessage { get; set; }
        public string ErrorStackTrace { get; set; }

        public static AgentErrorRecord Create(string source, Exception exception)
        {
            UserContextSnapshot userContext = UserContextSnapshot.Current();
            return new AgentErrorRecord
            {
                SyncId = Guid.NewGuid().ToString("N"),
                ErrorTime = DateTime.Now.ToString("yyyy.MM.dd. HH:mm:ss"),
                WindowsUser = userContext.UserName ?? string.Empty,
                SubDepartmentId = userContext.SubDepartmentId,
                RevitVersion = ModuleData.RevitVersion,
                Source = source ?? string.Empty,
                ErrorType = Truncate(exception == null ? string.Empty : exception.GetType().FullName),
                ErrorMessage = Truncate(exception == null ? string.Empty : exception.Message ?? string.Empty),
                ErrorStackTrace = Truncate(exception == null ? string.Empty : exception.ToString())
            };
        }

        private static string Truncate(string value)
        {
            if (string.IsNullOrEmpty(value) || value.Length <= MaxTextLength)
                return value ?? string.Empty;

            return value.Substring(0, MaxTextLength);
        }
    }
}