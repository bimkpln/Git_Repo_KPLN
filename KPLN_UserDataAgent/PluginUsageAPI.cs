using System;

namespace KPLN_UserDataAgent
{
    public static class PluginUsageApi
    {
        public static IDisposable BeginPluginExecution(string tabName, string buttonName)
        {
            return Module.BeginPluginExecution(tabName, buttonName);
        }

        public static IDisposable BeginPluginExecution(string buttonName)
        {
            return Module.BeginPluginExecution(null, buttonName);
        }
    }
}