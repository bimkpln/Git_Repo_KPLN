using Autodesk.Revit.DB;
using KPLN_Clashes_Ribbon.Core;
using KPLN_Library_ConfigWorker;
using System;

namespace KPLN_Clashes_Ribbon.Services
{
    internal static class ZoomSettingsConfigService
    {
        private const string ConfigName = "Clashes";

        public static ZoomSettings LoadOrCreateDefault(Document doc)
        {
            if (TryReadLocal(out ZoomSettings localModel))
                return localModel;

            return new ZoomSettings();
        }

        public static void Save(ZoomSettings model)
        {
            ConfigService.SaveConfig<ZoomSettings>(ConfigType.Local, model, ConfigName);
        }

        private static bool TryReadLocal(out ZoomSettings model)
        {
            model = null;

            try
            {
                object configObj = ConfigService.ReadConfigFile<ZoomSettings>(ConfigType.Local, ConfigName);
                model = configObj as ZoomSettings;
                return model != null;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
