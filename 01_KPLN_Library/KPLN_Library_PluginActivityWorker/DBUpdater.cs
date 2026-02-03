using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_SQLiteWorker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using System;
using System.Threading.Tasks;

namespace KPLN_Library_PluginActivityWorker
{
    public static class DBUpdater
    {
        public static async Task UpdatePluginActivityAsync_ByPluginNameAndModuleName(string pluginName, string moduleName)
        {
            try
            {
                await Task.Run(() =>
                {
                    UpdatePluginActivity_ByPluginNameAndDirName(pluginName, moduleName);
                });
            }
            catch (Exception ex)
            {
                HtmlOutput.Print(
                    $"Плагин успешно отработал. НО - не произошло записи активности плагина. Отправь ошибку разработчику: {ex.Message}",
                    MessageType.Warning);
            }
        }

        private static void UpdatePluginActivity_ByPluginNameAndDirName(string pluginName, string moduleName)
        {
            // Чистка имени плагина от переноса строк (в имени на вкладке это нужно, в БД - выглядит не очень)
            string clearPluginName = pluginName.Replace("\n", " ");


            // Для режима дебага и пользователей из вне - игнор
            if (DBMainService.CurrentDBUser.IsDebugMode || DBMainService.CurrentDBUser.IsExtraNet) return;


            // Обработка плагина по отделу и имени
            string currentDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
            DBPluginActivity currentPluginActivity = DBMainService.PluginActivityDbService.GetDBPluginActivity_ByModuleNameAndSubDep(clearPluginName, DBMainService.CurrentDBUser.SubDepartmentId);
            if (currentPluginActivity == null)
            {
                DBModule currentModule = DBMainService.ModuleDbService.GetDBModule_ByFiDirName(moduleName)
                    ?? throw new Exception($"Не удалось найти модуль по имени '{moduleName}'. Проверь БД, или ошибка при создании плагина.");

                currentPluginActivity = new DBPluginActivity()
                {
                    ModuleId = currentModule.Id,
                    PluginName = clearPluginName,
                    SubDepartmentId = DBMainService.CurrentDBUser.SubDepartmentId,
                    UsageCount = 1,
                    LastActivityDate = currentDate,
                };

                DBMainService.PluginActivityDbService.CreateDBPluginActivity(currentPluginActivity);
                return;
            }

            currentPluginActivity.UsageCount++;
            currentPluginActivity.LastActivityDate = currentDate;

            DBMainService.PluginActivityDbService.UpdatePluginActivity_ByPluginActivityAndSubDep(currentPluginActivity, DBMainService.CurrentDBUser.SubDepartmentId);
        }
    }
}
