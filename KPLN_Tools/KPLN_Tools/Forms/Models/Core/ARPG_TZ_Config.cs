using KPLN_Library_ConfigWorker.Core;
using System.Collections.ObjectModel;

namespace KPLN_Tools.Forms.Models.Core
{
    /// <summary>
    /// Обертка для конфига (для хранения в JSON)
    /// </summary>
    public sealed class ARPG_TZ_Config : IJsonSerializable
    {
        /// <summary>
        /// Ссылка на ARPG_MainData_TZ в окне
        /// </summary>
        public ARPG_TZ_MainData Config_ARPG_TZ_MainData { get; set; }

        /// <summary>
        /// Коллекция ARPG_FlatData_TZ в окне
        /// </summary>
        public ObservableCollection<ARPG_TZ_FlatData> Config_ARPG_TZ_FlatDataList { get; set; }

        public object ToJson()
        {
            return this;
        }
    }
}
