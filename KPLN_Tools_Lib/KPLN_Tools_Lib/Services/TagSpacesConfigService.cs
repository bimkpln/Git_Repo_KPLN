using Autodesk.Revit.DB;
using KPLN_Library_ConfigWorker;
using KPLN_Tools_Lib.Forms.Entities;
using System;

namespace KPLN_Tools_Lib.Services
{
    internal static class TagSpacesConfigService
    {
        internal const string ConfigName = "KPLN_Tools_Lib_TagSpaces";

        internal static TagSpacesConfig Load(Document document, out string warning)
        {
            warning = string.Empty;
            try
            {
                RequireSharedModel(document);
                return ConfigService.ReadConfigFile<TagSpacesConfig>(int.Parse(document.Application.VersionNumber),
                    document, ConfigType.Shared, ConfigName) as TagSpacesConfig;
            }
            catch (Exception ex)
            {
                warning = "Не удалось прочитать общий конфиг маркировки: " + ex.Message;
                return null;
            }
        }

        internal static string Save(Document document, TagSpacesConfig settings)
        {
            try
            {
                RequireSharedModel(document);
                ConfigService.SaveConfig<TagSpacesConfig>(int.Parse(document.Application.VersionNumber),
                    document, ConfigType.Shared, settings, ConfigName);
                return string.Empty;
            }
            catch (Exception ex)
            {
                return "Общий конфиг маркировки НЕ сохранён: " + ex.Message;
            }
        }

        private static void RequireSharedModel(Document document)
        {
            if (!document.IsWorkshared || document.IsDetached || document.GetWorksharingCentralModelPath() == null)
                throw new InvalidOperationException("Shared-конфиг доступен в модели, подключённой к хранилищу. "
                    + "Текущие настройки можно применить без сохранения.");
        }
    }
}
