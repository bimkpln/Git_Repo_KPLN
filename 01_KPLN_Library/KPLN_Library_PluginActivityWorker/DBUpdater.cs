using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using System;
using System.Threading.Tasks;

namespace KPLN_Library_PluginActivityWorker
{
    public static class DBUpdater
    {
        private static readonly UserDbService _userDbService = (UserDbService)new CreatorUserDbService().CreateService();
        private static readonly ModuleDbService _moduleDbService = (ModuleDbService)new CreatorModuleDbService().CreateService();
        private static readonly PluginActivityDbService _pluginActivityDbService = (PluginActivityDbService)new CreatorPluginActivityDbService().CreateService();

        private static DBUser _dBUser;

        public static async Task UpdatePluginActivityAsync_ByPluginNameAndDirName(string pluginName, string dllDirName)
        {
            try
            {
                await Task.Run(() =>
                {
                    UpdatePluginActivity_ByPluginNameAndDirName(pluginName, dllDirName);
                });
            }
            catch (Exception ex)
            {
                HtmlOutput.Print(
                    $"Плагин успешно отработал. НО - не произошло записи активности плагина. Отправь ошибку разработчику: {ex.Message}",
                    MessageType.Warning);
            }
        }

        private static void UpdatePluginActivity_ByPluginNameAndDirName(string pluginName, string dllDirName)
        {
            // Активирую данные по пользователю (позже переведи на KPLN_Loader, пока рано)
            if (_dBUser == null)
                _dBUser = _userDbService.GetCurrentDBUser();

            // Для режима дебага и пользователей из вне - игнор
            //if (_dBUser.IsDebugMode || _dBUser.IsExtraNet) return;

            string currentDate = DateTime.Now.ToString("yyyy-mm-dd HH:mm");
            // Обработка плагина по отделу и имени
            DBPluginActivity currentPluginActivity = _pluginActivityDbService.GetDBPluginActivity_ByModuleNameAndSubDep(pluginName, _dBUser.SubDepartmentId);
            if (currentPluginActivity == null)
            {
                DBModule currentModule = _moduleDbService.GetDBModule_ByFiDirName(dllDirName)
                    ?? throw new Exception($"Не удалось найти модуль по имени '{dllDirName}'. Проверь БД.");

                currentPluginActivity = new DBPluginActivity()
                {
                    ModuleId = currentModule.Id,
                    PluginName = pluginName,
                    SubDepartmentId = _dBUser.SubDepartmentId,
                    UsageCount = 1,
                    LastActivityDate = currentDate,
                };

                _pluginActivityDbService.CreateDBPluginActivity(currentPluginActivity);
                return;
            }

            currentPluginActivity.UsageCount++;

            _pluginActivityDbService.UpdatePluginActivity_ByPluginActivityAndSubDep(currentPluginActivity, _dBUser.SubDepartmentId);
        }
    }
}
