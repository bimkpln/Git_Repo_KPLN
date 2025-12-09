using Autodesk.Revit.DB;
using KPLN_Library_ConfigWorker;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UIFactory;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Tools.Common;
using KPLN_Tools.Common.OVVK_System;
using KPLN_Tools.ExecutableCommand;
using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace KPLN_Tools.Forms
{
    public partial class OV_DuctThicknessForm : Window
    {
        private readonly Document _doc;
        private readonly Element[] _elementsToSet;
        private readonly string _cofigName = "OV_DuctThickness";
        private readonly ConfigType _configType = ConfigType.Local;

        public OV_DuctThicknessForm(Document doc, Element[] elementsToSet)
        {
            _doc = doc;
            _elementsToSet = elementsToSet;

            ModelPath docModelPath = _doc.GetWorksharingCentralModelPath() ?? throw new Exception("Работает только с моделями из хранилища");
            string strDocModelPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(docModelPath);
            DBProject dBProject = DBWorkerService.CurrentProjectDbService.GetDBProject_ByRevitDocFileNameANDRVersion(strDocModelPath, ModuleData.RevitVersion);

            if (dBProject != null)
                _configType = ConfigType.Shared;

            #region Заполняю поля окна в зависимости от наличия файла конфига
            // Файл конфига присутсвует
            if (ConfigService.ReadConfigFile<DuctThicknessEntity>(ModuleData.RevitVersion, doc, _configType, _cofigName) is DuctThicknessEntity ductThicknessEntity)
                CurrentDuctThicknessEntity = ductThicknessEntity;
            else
            {
                CurrentDuctThicknessEntity = new DuctThicknessEntity()
                {
                    ParameterName = "КП_И_Толщина стенки",
                    PartOfInsulationName = "EI",
                    PartOfSystemName = "_Противодым._",
                };
            }
            #endregion

            InitializeComponent();

            this.DataContext = CurrentDuctThicknessEntity;
            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        public DuctThicknessEntity CurrentDuctThicknessEntity { get; private set; }

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
            if (e.Key == Key.Enter)
                StartBtn_Click(sender, e);
        }

        private void StartBtn_Click(object sender, RoutedEventArgs e)
        {
            Task.Run(() => { SaveConfig(); });

            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandDuctThickness_Start(CurrentDuctThicknessEntity, _elementsToSet));

            Close();
        }

        private void BtnParamSearch_Click(object sender, RoutedEventArgs e)
        {
            ElementSinglePick paramForm = SelectParameterFromRevit.CreateForm(this, _doc, _elementsToSet, StorageType.Double);
            paramForm.ShowDialog();

            if (paramForm.SelectedElement != null)
                CurrentDuctThicknessEntity.ParameterName = paramForm.SelectedElement.Name;
        }

        /// <summary>
        /// Сериализация и сохранение файла-конфигурации
        /// </summary>
        private void SaveConfig() => ConfigService.SaveConfig<DuctThicknessEntity>(ModuleData.RevitVersion, _doc, _configType, CurrentDuctThicknessEntity, _cofigName);
    }
}
