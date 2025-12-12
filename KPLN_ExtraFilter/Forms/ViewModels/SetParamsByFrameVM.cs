using Autodesk.Revit.DB;
using KPLN_ExtraFilter.ExecutableCommand;
using KPLN_ExtraFilter.Forms.Commands;
using KPLN_ExtraFilter.Forms.Entities;
using KPLN_ExtraFilter.Forms.Entities.SetParamsByFrame;
using KPLN_Library_ConfigWorker;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace KPLN_ExtraFilter.Forms.ViewModels
{
    public sealed class SetParamsByFrameVM
    {
        private readonly SetParamsByFrameForm _mainWindow;

        public SetParamsByFrameVM(SetParamsByFrameForm mainWindow, Document doc, IEnumerable<Element> userSelElems)
        {
            _mainWindow = mainWindow;
            CurrentSetParamsByFrameM = new SetParamsByFrameM(doc, userSelElems);
            CurrentSetParamsByFrameM.UpdateScriptANDCanRunANDUserHelp();

            // Чтение конфигурации последнего запуска
            object lastRunConfigObj = ConfigService.ReadConfigFile<SetParamsByFrameM_ParamM>(ModuleData.RevitVersion, doc, ConfigType.Memory);
            if (lastRunConfigObj != null && lastRunConfigObj is IEnumerable<SetParamsByFrameM_ParamM> paramMs)
            {
                foreach (SetParamsByFrameM_ParamM paramM in paramMs)
                {
                    if (paramM.ParamM_DocParameters == null || paramM.ParamM_SelectedParameter == null)
                        continue;
                    
                    var newParamM = new SetParamsByFrameM_ParamM(CurrentSetParamsByFrameM)
                    {
                        SearchParamText = paramM.SearchParamText,
                        ParamM_InputValue = paramM.ParamM_InputValue,
                    };
                    newParamM.RestoreSelectedParamById(paramM.ParamM_SelectedParameter.RevitParamIntId);
                    if (newParamM.ParamM_SelectedParameter == null)
                        newParamM.SearchParamText = string.Empty;

                    CurrentSetParamsByFrameM.ParamItems.Add(newParamM);
                }
            }


            // Установка команд
            AddNewParamCmd = new RelayCommand<object>(_ => AddNewParam());
            ClearingParamCmd = new RelayCommand<object>(_ => ClearingParam());
            RemoveParamCmd = new RelayCommand<SetParamsByFrameM_ParamM>(RemoveParam);
            SelectANDSetElemsParamCmd = new RelayCommand<object>(_ => SelectANDSetElemsParam());

            CloseWindowCmd = new RelayCommand<object>(CloseWindow);
        }

        public SetParamsByFrameM CurrentSetParamsByFrameM { get; set; }

        /// <summary>
        /// Комманда: Добаавить новый параметр
        /// </summary>
        public ICommand AddNewParamCmd { get; }

        /// <summary>
        /// Комманда: Очистить от всех параметров
        /// </summary>
        public ICommand ClearingParamCmd { get; }

        /// <summary>
        /// Комманда: Очистить выбранный параметр
        /// </summary>
        public ICommand RemoveParamCmd { get; }

        /// <summary>
        /// Комманда: Выбрать элементы в модели
        /// </summary>
        public ICommand SelectANDSetElemsParamCmd { get; }

        /// <summary>
        /// Комманда: Закрыть окно
        /// </summary>
        public ICommand CloseWindowCmd { get; }

        public void AddNewParam()
        {
            SetParamsByFrameM_ParamM defaultMI = new SetParamsByFrameM_ParamM(CurrentSetParamsByFrameM);
            CurrentSetParamsByFrameM.ParamItems.Add(defaultMI);

            CurrentSetParamsByFrameM.UpdateScriptANDCanRunANDUserHelp();
        }

        public void ClearingParam()
        {
            CurrentSetParamsByFrameM.ParamItems.Clear();

            CurrentSetParamsByFrameM.UpdateScriptANDCanRunANDUserHelp();
        }

        public void RemoveParam(SetParamsByFrameM_ParamM item)
        {
            string paramName = item.ParamM_SelectedParameter == null ? "Пустой параметр" : item.ParamM_SelectedParameter.RevitParamName;

            var td = MessageBox.Show(_mainWindow, $"Сейчас из будет удален параметр \"{paramName}\"", "Внимание", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (td == MessageBoxResult.Yes)
            {
                CurrentSetParamsByFrameM.ParamItems.Remove(item);

                CurrentSetParamsByFrameM.UpdateScriptANDCanRunANDUserHelp();
            }
        }

        public void SelectANDSetElemsParam()
        {
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new SetParamsByFrameExcCmd(CurrentSetParamsByFrameM));
#if Debug2020 || Revit2020
                _mainWindow.Close();
#endif
        }

        public void CloseWindow(object windObj)
        {
            if (windObj is Window window)
                window.Close();
        }
    }
}
