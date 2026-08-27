using KPLN_Library_ConfigWorker;
using KPLN_MEPBender.Forms.Entities;
using System;

namespace KPLN_MEPBender.Services.Config
{
    public sealed class MepBenderConfigService
    {
        private const string ConfigName = "KPLN_MEPBender";

        public MepBenderM LoadOrCreateDefault()
        {
            try
            {
                object configObj = ConfigService.ReadConfigFile<MepBenderM>(ConfigType.Local, ConfigName);
                if (configObj is MepBenderM model)
                    return model;
            }
            catch (Exception)
            {
                // Если файл отсутствует или повреждён, стартуем с дефолта и перезаписываем локальный конфиг.
            }

            MepBenderM defaultModel = new MepBenderM();
            Save(defaultModel);
            
            return defaultModel;
        }

        public void Save(MepBenderM model) => ConfigService.SaveConfig<MepBenderM>(ConfigType.Local, model, ConfigName);
    }
}
