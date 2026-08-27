using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_MEPBender.Common;
using KPLN_MEPBender.ExecutableCommand;
using KPLN_MEPBender.Forms.Commands;
using KPLN_MEPBender.Forms.Entities;
using KPLN_MEPBender.Services.Config;
using System.Windows;
using System.Windows.Input;

namespace KPLN_MEPBender.Forms.ViewModels
{
    public sealed class MepBenderVM
    {
        private readonly MepBenderForm _mainWindow;
        private readonly MepBenderConfigService _configService;

        public MepBenderVM(MepBenderForm mainWindow, UIApplication uiapp)
        {
            _mainWindow = mainWindow;
            _configService = new MepBenderConfigService();
            CurrentMepBenderM = _configService.LoadOrCreateDefault();

            PickModelObstaclesCmd = new RelayCommand<object>(_ => PickObstacles(MepBenderObstacleSelectionSource.Model));
            PickLinkedObstaclesCmd = new RelayCommand<object>(_ => PickObstacles(MepBenderObstacleSelectionSource.Link));
            ClearObstaclesCmd = new RelayCommand<object>(_ => CurrentMepBenderM.ClearObstacles());
            PickRoutesCmd = new RelayCommand<object>(_ => PickRoutes());
            ClearRoutesCmd = new RelayCommand<object>(_ => CurrentMepBenderM.ClearRoutes());
            RunBenderCmd = new RelayCommand<object>(_ => RunBender());
            CloseWindowCmd = new RelayCommand<object>(CloseWindow);
        }

        public MepBenderM CurrentMepBenderM { get; set; }

        public ICommand PickModelObstaclesCmd { get; }

        public ICommand PickLinkedObstaclesCmd { get; }

        public ICommand ClearObstaclesCmd { get; }

        public ICommand PickRoutesCmd { get; }

        public ICommand ClearRoutesCmd { get; }

        public ICommand RunBenderCmd { get; }

        public ICommand CloseWindowCmd { get; }

        private void PickObstacles(MepBenderObstacleSelectionSource source)
        {
            CurrentMepBenderM.SetStatus(null, source == MepBenderObstacleSelectionSource.Link
                ? "Выбери элементы из связанной модели, которые трасса должна обогнуть."
                : "Выбери элементы из текущей модели, которые трасса должна обогнуть.");

            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new SelectObstaclesExcCmd(CurrentMepBenderM, source));
        }

        private void PickRoutes()
        {
            CurrentMepBenderM.SetStatus(null, "Выбери трубы, воздуховоды или кабельные лотки для изменения.");
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new SelectRouteElementsExcCmd(CurrentMepBenderM));
        }

        private void RunBender()
        {
            if (!CurrentMepBenderM.CanRun)
            {
                MessageBox.Show(
                    _mainWindow,
                    "Для запуска нужны огибаемые элементы, MEP-трассы, направление, смещение и угол из списка.",
                    "Внимание",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new BendRoutesExcCmd(CurrentMepBenderM));
            _configService.Save(CurrentMepBenderM);
        }

        private void CloseWindow(object windObj)
        {
            if (windObj is Window window)
                window.Close();
        }
    }
}
