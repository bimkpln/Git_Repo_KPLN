using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_ConfigWorker;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_Forms.UIFactory;
using KPLN_Tools_OVVK.ExecutableCommand;
using KPLN_Tools_OVVK.Common.OVVK_System;
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

        private readonly Window _mainWindow;

        public OVVK_SystemManagerVM(Window mainWindow, Document doc)
        {
            _mainWindow = mainWindow;
            CurrentOVVK_SystemManagerM = new OVVK_SystemManagerM(doc);

            SelectParameterCmd = new RelayCommand<object>(_ => QueueAction(SelectParameter), _ => CurrentOVVK_SystemManagerM.CanSelectParameter);
            RefreshSystemsCmd = new RelayCommand<object>(_ => QueueAction(RefreshSystems), _ => CurrentOVVK_SystemManagerM.CanRefresh);
            RefreshSystemNamesCmd = new RelayCommand<object>(_ => QueueAction(app =>
                new OVVK_SystemRefreshExcCmd(CurrentOVVK_SystemManagerM, _mainWindow).Execute(app)),
                _ => CurrentOVVK_SystemManagerM.CanRefresh);
            FillSystemNameCmd = new RelayCommand<object>(_ => QueueAction(FillSystemName), _ => CurrentOVVK_SystemManagerM.CanFillSystemName);
            SelectSystemCmd = new RelayCommand<object>(_ => SelectSystem(), _ => CurrentOVVK_SystemManagerM.CanSelectSystem);
            MergeSystemsCmd = new RelayCommand<object>(_ => QueueAction(MergeSystems), _ => CurrentOVVK_SystemManagerM.CanMergeSystems);
            MergeSelectedSystemsCmd = new RelayCommand<object>(_ => QueueAction(MergeSelectedSystems), _ => CurrentOVVK_SystemManagerM.CanMergeSelectedSystems);
            CreateViewsCmd = new RelayCommand<object>(_ =>
            {
                string[] selectedNames = CurrentOVVK_SystemManagerM.SelectedSystemNames;
                QueueAction(app => CreateViews(app, selectedNames));
            }, _ => CurrentOVVK_SystemManagerM.CanCreateViews);
            OrientViewsCmd = new RelayCommand<object>(_ => QueueAction(app =>
                new OVVK_OrientViewsExcCmd(CurrentOVVK_SystemManagerM, _mainWindow).Execute(app)),
                _ => CurrentOVVK_SystemManagerM.CanOrientViews);
            CloseWindowCmd = new RelayCommand<object>(_ => _mainWindow.Close());

            CurrentOVVK_SystemManagerM.PropertyChanged += OnModelPropertyChanged;
            LoadConfig();
        }


        public OVVK_SystemManagerM CurrentOVVK_SystemManagerM { get; }
        public ICommand SelectParameterCmd { get; }
        public ICommand RefreshSystemsCmd { get; }
        public ICommand RefreshSystemNamesCmd { get; }
        public ICommand FillSystemNameCmd { get; }
        public ICommand SelectSystemCmd { get; }
        public ICommand MergeSystemsCmd { get; }
        public ICommand MergeSelectedSystemsCmd { get; }
        public ICommand CreateViewsCmd { get; }
        public ICommand OrientViewsCmd { get; }
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


        private void SelectSystem()
        {
            // Запоминаем выбранные строки до выполнения очереди Revit.
            OVVK_SystemManagerRowM[] selectedRows = CurrentOVVK_SystemManagerM.SelectedSystemRows;
            QueueAction(app =>
            {
                CurrentOVVK_SystemManagerM.Refresh();
                app.ActiveUIDocument.Selection.SetElementIds(selectedRows
                    .SelectMany(row => CurrentOVVK_SystemManagerM.GetSystemElementIds(row)).Distinct().ToArray());
            });
        }


        // Следующие операции вызываются очередью команд в контексте Revit API.
        private void SelectParameter(UIApplication app)
        {
            CurrentOVVK_SystemManagerM.Refresh();
            if (CurrentOVVK_SystemManagerM.ElementColl.Count == 0)
                return;

            ElementSinglePick parameterForm = SelectParameterFromRevit.CreateForm(
                _mainWindow,
                CurrentOVVK_SystemManagerM.CurrentDoc,
                CurrentOVVK_SystemManagerM.ElementColl,
                StorageType.String,
                true,
                true);
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

        private void FillSystemName(UIApplication app)
        {
            if (!CurrentOVVK_SystemManagerM.ValidateSystemAbbreviations())
                return;

            string parameterName = CurrentOVVK_SystemManagerM.ParameterName;
            if (string.IsNullOrWhiteSpace(parameterName))
                return;

            var currentViewButton = new ButtonToRunEntity("На текущем виде", "Заполнить элементы ОВВК на открытом виде с учётом фильтра менеджера.");
            var wholeModelButton = new ButtonToRunEntity("На всю модель", "Заполнить элементы ОВВК во всей модели с учётом фильтра менеджера.");
            var scopeForm = new ButtonToRun($"Где заполнить «{parameterName}»?", new[] { currentViewButton, wholeModelButton })
            {
                Owner = _mainWindow,
                Title = "KPLN: Область заполнения параметра",
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
            };
            if (scopeForm.ShowDialog() != true || scopeForm.SelectedButton == null)
                return;


            ElementId viewId = ReferenceEquals(scopeForm.SelectedButton, currentViewButton)
                ? CurrentOVVK_SystemManagerM.CurrentDoc.ActiveView.Id
                : null;
            new OVVK_SystemNameFillExcCmd(CurrentOVVK_SystemManagerM, parameterName,
                viewId, _mainWindow).Execute(app);
        }

        private void MergeSystems(UIApplication app)
        {
            CurrentOVVK_SystemManagerM.Refresh();
            if (CurrentOVVK_SystemManagerM.HasParameterErrors || CurrentOVVK_SystemManagerM.HasTypeParameters
                || CurrentOVVK_SystemManagerM.SystemSumParameters.Count < 2)
                return;

            var mergeForm = new OVVK_MergeSystemsForm(CurrentOVVK_SystemManagerM) { Owner = _mainWindow };
            if (mergeForm.ShowDialog() != true)
                return;

            new ExcCmd_SystemManager_MergeSystem(CurrentOVVK_SystemManagerM, mergeForm.SysDataToMerge.ToArray()).Execute(app);
            CurrentOVVK_SystemManagerM.Refresh();
        }

        private void MergeSelectedSystems(UIApplication app)
        {
            var model = CurrentOVVK_SystemManagerM;
            ElementId[] ids = app.ActiveUIDocument.Selection.GetElementIds().ToArray();
            if (!OVVK_MergeSystem.IsAllowedCount(ids.Length))
            {
                model.SetOperationMessage($"Выдели от {OVVK_MergeSystem.MinSystemCount} до {OVVK_MergeSystem.MaxSystemCount} элементов ОВВК.");
                return;
            }

            model.Refresh();
            if (model.HasParameterErrors || model.HasTypeParameters || string.IsNullOrWhiteSpace(model.ParameterName))
            {
                model.SetOperationMessage("Объединение недоступно: выбери параметр экземпляра и исправь ошибки его значений.");
                return;
            }

            var allowed = model.ElementColl.ToDictionary(element => element.Id);
            var names = new System.Collections.Generic.List<string>();
            foreach (ElementId id in ids)
            {
                if (!allowed.TryGetValue(id, out Element element))
                {
                    model.SetOperationMessage("Выделение содержит элементы вне списка ОВВК менеджера. Выбери элементы нужных систем.");
                    return;
                }
                string name = element.LookupParameter(model.ParameterName)?.AsString();
                if (string.IsNullOrWhiteSpace(name) || !model.SystemSumParameters.Contains(name))
                {
                    model.SetOperationMessage($"У элемента {id} не определена корректная система в параметре «{model.ParameterName}».");
                    return;
                }
                if (!names.Contains(name))
                    names.Add(name);
            }
            if (!OVVK_MergeSystem.IsAllowedCount(names.Count))
            {
                model.SetOperationMessage("Выбранные элементы относятся к одной системе. Выбери элементы минимум двух разных систем.");
                return;
            }

            var group = new OVVK_MergeSystem(model.SystemSumParameters.ToArray());
            group.SetSystemNames(names);
            new ExcCmd_SystemManager_MergeSystem(model, new[] { group }).Execute(app);
            model.Refresh();
        }

        private void CreateViews(UIApplication app, string[] selectedNames)
        {
            CurrentOVVK_SystemManagerM.Refresh();
            if (CurrentOVVK_SystemManagerM.HasParameterErrors || CurrentOVVK_SystemManagerM.HasSystemValidationErrors
                || CurrentOVVK_SystemManagerM.HasDuplicateSystemNames)
            {
                CurrentOVVK_SystemManagerM.SetOperationMessage("Создание 3D-видов недоступно: исправь все критические ошибки, отмеченные красным.");
                return;
            }

            string[] systemNames = selectedNames.Where(name => CurrentOVVK_SystemManagerM.SystemSumParameters.Contains(name)).ToArray();
            if (systemNames.Length == 0)
            {
                CurrentOVVK_SystemManagerM.SetOperationMessage("Выбранные системы больше не доступны. Выбери обычные строки систем в списке.");
                return;
            }

            View activeView = CurrentOVVK_SystemManagerM.CurrentDoc.ActiveView;
            if (!(activeView is View3D) || activeView.IsTemplate || !activeView.CanViewBeDuplicated(ViewDuplicateOption.Duplicate))
            {
                MessageBox.Show(_mainWindow, "Открой 3D-вид, на основе которого нужно создать схемы систем.",
                    "KPLN: Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }


            // Проверяем управление фильтрами до открытия выбора систем.
            View viewTemplate = CurrentOVVK_SystemManagerM.CurrentDoc.GetElement(activeView.ViewTemplateId) as View;
            if (!ExcCmd_SystemManager_ViewCreator.ValidateTemplateFilters(viewTemplate))
                return;


            Result result = new ExcCmd_SystemManager_ViewCreator(CurrentOVVK_SystemManagerM,
                systemNames, dataIsCurrent: true).Execute(app);
            if (result == Result.Succeeded)
                _mainWindow.Close();
        }

        private void LoadConfig()
        {
            try
            {
                Document doc = CurrentOVVK_SystemManagerM.CurrentDoc;
                OVVK_SystemManagerM saved = null;
                SystemProjectConfig.Access(doc, root =>
                {
                    saved = root["Settings"].ToObject<OVVK_SystemManagerM>();
                    return false;
                });
                if (saved == null)
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
                Document doc = CurrentOVVK_SystemManagerM.CurrentDoc;
                SystemProjectConfig.Access(doc, root =>
                {
                    root["Settings"] = Newtonsoft.Json.Linq.JObject.FromObject(CurrentOVVK_SystemManagerM.ToJson());
                    return true;
                });
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
