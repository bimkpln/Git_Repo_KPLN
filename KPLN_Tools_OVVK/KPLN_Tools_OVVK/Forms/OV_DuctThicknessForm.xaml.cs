using Autodesk.Revit.DB;
using KPLN_Library_ConfigWorker;
using KPLN_Library_DBWorker;
using KPLN_Library_DBWorker.Core;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UIFactory;
using KPLN_Tools_OVVK.Common.OVVK_System;
using KPLN_Tools_OVVK.ExecutableCommand;
using System.Windows;
using System.Windows.Input;

namespace KPLN_Tools_OVVK.Forms
{
    public partial class OV_DuctThicknessForm : Window
    {
        private readonly Document _doc;
        private readonly Element[] _elementsToSet;
        private readonly ConfigType _configType = ConfigType.Local;

        public OV_DuctThicknessForm(Document doc, Element[] elementsToSet)
        {
            _doc = doc;
            _elementsToSet = elementsToSet;

            ModelPath docModelPath = _doc.IsWorkshared ? _doc.GetWorksharingCentralModelPath() : null;
            if (docModelPath != null)
            {
                string strDocModelPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(docModelPath);
                DBProject dBProject = SQLiteMainService.SQLitePrjServiceInst.GetDBProject_ByRevitDocFileNameANDRVersion(strDocModelPath, ModuleData.RevitVersion);
                if (dBProject != null)
                    _configType = ConfigType.Shared;
            }

            #region Заполняю поля окна в зависимости от наличия файла конфига
            CurrentDuctThicknessEntity = ConfigService.ReadConfigFile<DuctThicknessEntity>(
                ModuleData.RevitVersion, doc, _configType, DuctThicknessEntity.ConfigName) as DuctThicknessEntity;
            if (CurrentDuctThicknessEntity == null && _configType == ConfigType.Shared)
                CurrentDuctThicknessEntity = ConfigService.ReadConfigFile<DuctThicknessEntity>(
                    ModuleData.RevitVersion, doc, ConfigType.Local, DuctThicknessEntity.ConfigName) as DuctThicknessEntity;
            if (CurrentDuctThicknessEntity == null)
                CurrentDuctThicknessEntity = new DuctThicknessEntity();
            #endregion

            CurrentDuctThicknessEntity.UseProtectionParameters = DuctProtectionParameters.AreAvailable(doc);

            InitializeComponent();

            LegacySettingsPanel.Visibility = CurrentDuctThicknessEntity.UseProtectionParameters
                ? System.Windows.Visibility.Collapsed : System.Windows.Visibility.Visible;
            ProtectionParametersHelp.Visibility = CurrentDuctThicknessEntity.UseProtectionParameters
                ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;

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
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(
                new CommandDuctThickness_Start(CurrentDuctThicknessEntity, _elementsToSet, _configType));

            Close();
        }

        private void BtnParamSearch_Click(object sender, RoutedEventArgs e)
        {
            ElementSinglePick paramForm = SelectParameterFromRevit.CreateForm(this, _doc, _elementsToSet, StorageType.Double);
            paramForm.ShowDialog();

            if (paramForm.SelectedElement != null)
                CurrentDuctThicknessEntity.ParameterName = paramForm.SelectedElement.Name;
        }
    }
}
