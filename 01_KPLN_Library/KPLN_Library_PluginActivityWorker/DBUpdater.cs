using KPLN_Library_DBWorker;
using KPLN_Library_DBWorker.Core;
using KPLN_Library_Forms.UI.HtmlWindow;
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
            if (SQLiteMainService.CurrentDBUser.IsDebugMode || SQLiteMainService.CurrentDBUser.IsExtraNet) return;


            // Обработка плагина по отделу и имени
            string currentDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
            DBPluginActivity currentPluginActivity = SQLiteMainService
                .SQLitePluginActivityServiceInst
                .GetDBPluginActivity_ByModuleNameAndSubDep(clearPluginName, SQLiteMainService.CurrentDBUser.SubDepartmentId);
            if (currentPluginActivity == null)
            {
                DBModule currentModule = SQLiteMainService.SQLiteModuleServiceInst.GetDBModule_ByFiDirName(moduleName)
                    ?? throw new Exception($"Не удалось найти модуль по имени '{moduleName}'. Проверь БД, или ошибка при создании плагина.");

                currentPluginActivity = new DBPluginActivity()
                {
                    ModuleId = currentModule.Id,
                    PluginName = clearPluginName,
                    SubDepartmentId = SQLiteMainService.CurrentDBUser.SubDepartmentId,
                    UsageCount = 1,
                    LastActivityDate = currentDate,
                };

                SQLiteMainService.SQLitePluginActivityServiceInst.CreateDBPluginActivity(currentPluginActivity);
                return;
            }

            currentPluginActivity.UsageCount++;
            currentPluginActivity.LastActivityDate = currentDate;

            SQLiteMainService.SQLitePluginActivityServiceInst.UpdatePluginActivity_ByPluginActivityAndSubDep(currentPluginActivity, SQLiteMainService.CurrentDBUser.SubDepartmentId);
        }
    }
}
