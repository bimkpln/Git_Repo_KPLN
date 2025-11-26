using Autodesk.Revit.DB;
using KPLN_ExtraFilter.ExecutableCommand;
using KPLN_ExtraFilter.Forms.Commands;
using KPLN_ExtraFilter.Forms.Entities;
using KPLN_Library_ConfigWorker;
using System.Windows.Input;

namespace KPLN_ExtraFilter.Forms.ViewModels
{
    public sealed class SelectionByClickVM
    {
        public SelectionByClickVM(Document doc)
        {
            // Чтение конфигурации последнего запуска
            object lastRunConfigObj = ConfigService.ReadConfigFile<SelectionByClickM>(ModuleData.RevitVersion, doc, ConfigType.Memory);
            if (lastRunConfigObj != null && lastRunConfigObj is SelectionByClickM model)
                CurrentSelectionByClickM = model;
            else
                CurrentSelectionByClickM = new SelectionByClickM(doc);


            RunSelectionCmd = new RelayCommand<object>(_ => RunSelection());
            DropSelectionCmd = new RelayCommand<object>(_ => DropSelection());
        }


        public SelectionByClickM CurrentSelectionByClickM { get; set; }

        /// <summary>
        /// Комманда: Запуск выбора
        /// </summary>
        public ICommand RunSelectionCmd { get; }

        /// <summary>
        /// Комманда: Сброс выбора
        /// </summary>
        public ICommand DropSelectionCmd { get; }

        public void RunSelection()
        {
            // Запись конфигурации последнего запуска
            ConfigService.SaveConfig<SelectionByClickM>(ModuleData.RevitVersion, CurrentSelectionByClickM.UserSelDoc, ConfigType.Memory, CurrentSelectionByClickM);

            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new SelectionByClickExcCmd(CurrentSelectionByClickM));
        }

        public void DropSelection() => CurrentSelectionByClickM.DropToDefault();
    }
}
