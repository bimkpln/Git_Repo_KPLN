using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_ConfigWorker;
using KPLN_Library_ConfigWorker.Core;
using KPLN_Library_DBWorker;
using KPLN_Library_DBWorker.Core;
using KPLN_OpeningHoleManager.Forms.MVVMCore_MainMenu;
using System;

namespace KPLN_OpeningHoleManager.Core
{
    public sealed class MainConfig : IJsonSerializable
    {
        private static readonly string _cofigName = "AR_OHEManagerConfig";


        public MainConfig()
        {
        }

        /// <summary>
        /// Значение расширение отверстия при расстановке
        /// </summary>
        public double OpenHoleExpandedValue { get; set; }

        public int OpenHoleZRoundStepMm { get; set; } = 10;

        public OpeningTypeAssignment[] OpeningTypeAssignments { get; set; } = new OpeningTypeAssignment[0];

        /// <summary>
        /// АР: Значение минимальной ширины отверстия при расстановке
        /// </summary>
        public double AR_OpenHoleMinWidthValue { get; set; }

        /// <summary>
        /// АР: Значение минимальной высоты отверстия при расстановке
        /// </summary>
        public double AR_OpenHoleMinHeightValue { get; set; }

        /// <summary>
        /// АР: Значение минимального расстояни между отверстиями для объединения (UV)
        /// </summary>
        public double AR_OpenHoleMinDistanceValue { get; set; }

        /// <summary>
        /// КР: Значение минимальной ширины отверстия при расстановке
        /// </summary>
        public double KR_OpenHoleMinWidthValue { get; set; }

        /// <summary>
        /// КР: Значение минимальной высоты отверстия при расстановке
        /// </summary>
        public double KR_OpenHoleMinHeightValue { get; set; }

        /// <summary>
        /// КР: Значение минимального расстояни между отверстиями для объединения (UV)
        /// </summary>
        public double KR_OpenHoleMinDistanceValue { get; set; }


        public static MainConfig GetData_FromConfig(DBProject dBProject)
        {
            if (Module.CurrentUIApplication?.ActiveUIDocument == null) return null;

            UIDocument uidoc = Module.CurrentUIApplication.ActiveUIDocument;
            Document doc = uidoc.Document;

            ConfigType configType = doc.IsWorkshared && dBProject != null ? ConfigType.Shared : ConfigType.Local;

            object obj = ConfigService.ReadConfigFile<MainConfig>(ModuleData.RevitVersion, doc, configType, _cofigName);
            if (obj is MainConfig configItem)
            {
                return new MainConfig()
                {
                    OpenHoleExpandedValue = configItem.OpenHoleExpandedValue,
                    OpenHoleZRoundStepMm = configItem.OpenHoleZRoundStepMm == 5 ? 5 : 10,
                    OpeningTypeAssignments = configItem.OpeningTypeAssignments,

                    AR_OpenHoleMinDistanceValue = configItem.AR_OpenHoleMinDistanceValue,
                    AR_OpenHoleMinHeightValue = configItem.AR_OpenHoleMinHeightValue,
                    AR_OpenHoleMinWidthValue = configItem.AR_OpenHoleMinWidthValue,

                    KR_OpenHoleMinDistanceValue = configItem.KR_OpenHoleMinDistanceValue,
                    KR_OpenHoleMinHeightValue = configItem.KR_OpenHoleMinHeightValue,
                    KR_OpenHoleMinWidthValue = configItem.KR_OpenHoleMinWidthValue,
                };
            }

            return null;
        }

        public static void SetData_ToConfig(MainViewModel vm)
        {
            if (Module.CurrentUIApplication?.ActiveUIDocument == null) return;

            UIDocument uidoc = Module.CurrentUIApplication.ActiveUIDocument;
            Document doc = uidoc.Document;

            ModelPath docModelPath = doc.IsWorkshared ? doc.GetWorksharingCentralModelPath() : null;
            string strDocModelPath = docModelPath == null ? doc.PathName : ModelPathUtils.ConvertModelPathToUserVisiblePath(docModelPath);
            DBProject dBProject = SQLiteMainService.SQLitePrjServiceInst.GetDBProject_ByRevitDocFileNameANDRVersion(strDocModelPath, ModuleData.RevitVersion);
            ConfigType configType = docModelPath != null && dBProject != null ? ConfigType.Shared : ConfigType.Local;

            MainConfig mainConfig = new MainConfig()
            {
                OpenHoleExpandedValue = vm.OpenHoleExpandedValue,
                OpenHoleZRoundStepMm = vm.OpenHoleZRoundStepMm,
                OpeningTypeAssignments = vm.TypeMapping.Export(),

                AR_OpenHoleMinDistanceValue = vm.AR_OpenHoleMinDistanceValue,
                AR_OpenHoleMinHeightValue = vm.AR_OpenHoleMinHeightValue,
                AR_OpenHoleMinWidthValue = vm.AR_OpenHoleMinWidthValue,

                KR_OpenHoleMinDistanceValue = vm.KR_OpenHoleMinDistanceValue,
                KR_OpenHoleMinHeightValue = vm.KR_OpenHoleMinHeightValue,
                KR_OpenHoleMinWidthValue = vm.KR_OpenHoleMinWidthValue,
            };

            ConfigService.SaveConfig<MainConfig>(ModuleData.RevitVersion, doc, configType, mainConfig, _cofigName);
        }

        public object ToJson()
        {
            return this;
        }
    }
}
