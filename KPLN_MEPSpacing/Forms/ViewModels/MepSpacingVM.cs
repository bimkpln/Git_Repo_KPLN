using Autodesk.Revit.UI;
using KPLN_MEPSpacing.ExecutableCommand;
using KPLN_MEPSpacing.Forms.Commands;
using KPLN_MEPSpacing.Forms.Entities;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;

namespace KPLN_MEPSpacing.Forms.ViewModels
{
    public sealed class MepSpacingVM
    {
        private readonly MepSpacingForm _mainWindow;

        public MepSpacingVM(MepSpacingForm mainWindow, UIApplication uiapp, IEnumerable<Autodesk.Revit.DB.ElementId> selectedElementIds)
        {
            _mainWindow = mainWindow;
            CurrentMepSpacingM = new MepSpacingM(uiapp, selectedElementIds);

            PickSpacingElementsCmd = new RelayCommand<object>(_ => PickSpacingElements());
            PickBaseElementsCmd = new RelayCommand<object>(_ => PickBaseElements());
            RunSpacingCmd = new RelayCommand<object>(_ => RunSpacing());
            CloseWindowCmd = new RelayCommand<object>(CloseWindow);
        }

        public MepSpacingM CurrentMepSpacingM { get; }

        public ICommand PickSpacingElementsCmd { get; }

        public ICommand PickBaseElementsCmd { get; }

        public ICommand RunSpacingCmd { get; }

        public ICommand CloseWindowCmd { get; }

        public void PickSpacingElements() =>
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new PickSpacingElementsExcCmd(CurrentMepSpacingM, _mainWindow));

        public void PickBaseElements() =>
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new PickBaseElementsExcCmd(CurrentMepSpacingM, _mainWindow));

        public void RunSpacing() =>
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new ApplySpacingExcCmd(CurrentMepSpacingM));

        public void CloseWindow(object windObj)
        {
            if (windObj is Window window)
                window.Close();
        }
    }
}
