using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_ConfigWorker;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_Forms.UIFactory;
using KPLN_Tools_OVVK.ExecutableCommand;
using KPLN_Tools_OVVK.Forms.Entities;
using KPLN_Tools_OVVK.Forms.Models.Core;
using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace KPLN_Tools_OVVK.Forms.ViewModels
{
    public sealed class OVVK_SystemManagerVM
    {
        private const string ConfigName = "OVVK_SystemManager";
        private readonly Window _mainWindow;

        public OVVK_SystemManagerVM(Window mainWindow, Document doc)
        {
            _mainWindow = mainWindow;
            CurrentOVVK_SystemManagerM = new OVVK_SystemManagerM(doc);

            SelectParameterCmd = new RelayCommand<object>(_ => QueueAction(SelectParameter), _ => CurrentOVVK_SystemManagerM.CanSelectParameter);
            RefreshSystemsCmd = new RelayCommand<object>(_ => QueueAction(RefreshSystems), _ => CurrentOVVK_SystemManagerM.CanRefresh);
            MergeSystemsCmd = new RelayCommand<object>(_ => QueueAction(MergeSystems), _ => CurrentOVVK_SystemManagerM.CanMergeSystems);
            CreateViewsCmd = new RelayCommand<object>(_ => QueueAction(CreateViews), _ => CurrentOVVK_SystemManagerM.CanCreateViews);
            SelectWarningsCmd = new RelayCommand<object>(_ => QueueAction(SelectWarnings), _ => CurrentOVVK_SystemManagerM.CanSelectWarnings);
            CloseWindowCmd = new RelayCommand<object>(_ => _mainWindow.Close());

            CurrentOVVK_SystemManagerM.PropertyChanged += OnModelPropertyChanged;
            LoadConfig();
        }


        public OVVK_SystemManagerM CurrentOVVK_SystemManagerM { get; }
        public ICommand SelectParameterCmd { get; }
        public ICommand RefreshSystemsCmd { get; }
        public ICommand MergeSystemsCmd { get; }
        public ICommand CreateViewsCmd { get; }
        public ICommand SelectWarningsCmd { get; }
        public ICommand CloseWindowCmd { get; }


        internal void Release()
        {
            CurrentOVVK_SystemManagerM.IsClosed = true;
            CurrentOVVK_SystemManagerM.PropertyChanged -= OnModelPropertyChanged;
            SaveConfig();
        }

        private void QueueAction(Action<UIApplication> action)
        {
            if (!CurrentOVVK_SystemManagerM.CanRefresh)
                return;

            CurrentOVVK_SystemManagerM.IsBusy = true;
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(
                new OVVK_SystemManagerExcCmd(CurrentOVVK_SystemManagerM, action));
        }


        // Следующие операции вызываются очередью команд в контексте Revit API.
        private void SelectParameter(UIApplication app)
        {
            CurrentOVVK_SystemManagerM.Refresh();
            if (CurrentOVVK_SystemManagerM.ElementColl.Count == 0)
                return;

            ElementSinglePick parameterForm = SelectParameterFromRevit.CreateForm(
                _mainWindow, CurrentOVVK_SystemManagerM.CurrentDoc,
                CurrentOVVK_SystemManagerM.ElementColl, StorageType.String);
            if (parameterForm.ShowDialog() != true || parameterForm.SelectedElement == null)
                return;

            CurrentOVVK_SystemManagerM.ParameterName = parameterForm.SelectedElement.Name;
            CurrentOVVK_SystemManagerM.UpdateSystemParamData();
            SaveConfig();
        }

        private void RefreshSystems(UIApplication app)
        {
            new ExcCmd_SystemManager_VMUpdater(CurrentOVVK_SystemManagerM).Execute(app);
        }

        private void MergeSystems(UIApplication app)
        {
            CurrentOVVK_SystemManagerM.Refresh();
            if (CurrentOVVK_SystemManagerM.HasParameterErrors || CurrentOVVK_SystemManagerM.SystemSumParameters.Count < 2)
                return;

            var mergeForm = new OVVK_MergeSystemsForm(CurrentOVVK_SystemManagerM) { Owner = _mainWindow };
            if (mergeForm.ShowDialog() != true)
                return;

            new ExcCmd_SystemManager_MergeSystem(CurrentOVVK_SystemManagerM, mergeForm.SysDataToMerge.ToArray()).Execute(app);
            CurrentOVVK_SystemManagerM.Refresh();
        }

        private void CreateViews(UIApplication app)
        {
            CurrentOVVK_SystemManagerM.Refresh();
            if (CurrentOVVK_SystemManagerM.HasParameterErrors || CurrentOVVK_SystemManagerM.SystemSumParameters.Count == 0)
                return;

            View activeView = CurrentOVVK_SystemManagerM.CurrentDoc.ActiveView;
            if (!(activeView is View3D) || activeView.IsTemplate || !activeView.CanViewBeDuplicated(ViewDuplicateOption.Duplicate))
            {
                MessageBox.Show(_mainWindow, "Открой 3D-вид, на основе которого нужно создать схемы систем.",
                    "KPLN: Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }


            var systemsForm = new ElementMultiPick(_mainWindow,
                CurrentOVVK_SystemManagerM.SystemSumParameters.Select(name => new ElementEntity(name)));
            if (systemsForm.ShowDialog() != true || !systemsForm.SelectedElements.Any())
                return;

            Result result = new ExcCmd_SystemManager_ViewCreator(CurrentOVVK_SystemManagerM,
                systemsForm.SelectedElements.Select(element => element.Name).ToArray()).Execute(app);
            if (result == Result.Succeeded)
                _mainWindow.Close();
        }

        private void SelectWarnings(UIApplication app)
        {
            CurrentOVVK_SystemManagerM.Refresh();
            Element[] warningElements = CurrentOVVK_SystemManagerM.WarningsElementColl.ToArray();
            if (warningElements.Length == 0)
                return;

            new ExcCmd_ShowElement(warningElements).Execute(app);
            _mainWindow.Close();
        }


        private void LoadConfig()
        {
            try
            {
                var saved = ConfigService.ReadConfigFile<OVVK_SystemManagerM>(ConfigType.Local, ConfigName) as OVVK_SystemManagerM;
                if (saved == null || !string.Equals(saved.DocumentKey, CurrentOVVK_SystemManagerM.DocumentKey, StringComparison.OrdinalIgnoreCase))
                    return;

                CurrentOVVK_SystemManagerM.ParameterName = saved.ParameterName;
                CurrentOVVK_SystemManagerM.SysNameSeparator = saved.SysNameSeparator;
                CurrentOVVK_SystemManagerM.UpdateSystemParamData();
            }
            catch (Exception ex)
            {
                CurrentOVVK_SystemManagerM.SetOperationMessage("Не удалось прочитать настройки. Выбери параметр заново.");
                HtmlOutput.PrintError(ex);
            }
        }

        private void SaveConfig()
        {
            if (string.IsNullOrWhiteSpace(CurrentOVVK_SystemManagerM.ParameterName))
                return;

            try
            {
                ConfigService.SaveConfig<OVVK_SystemManagerM>(ConfigType.Local, CurrentOVVK_SystemManagerM, ConfigName);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось сохранить настройки: " + ex.Message,
                    "KPLN: Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void OnModelPropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            CommandManager.InvalidateRequerySuggested();
        }
    }
}
