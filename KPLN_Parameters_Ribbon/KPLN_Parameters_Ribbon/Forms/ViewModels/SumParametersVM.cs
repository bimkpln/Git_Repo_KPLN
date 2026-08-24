using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Parameters_Ribbon.Command;
using KPLN_Parameters_Ribbon.Forms.Commands;
using KPLN_Parameters_Ribbon.Forms.Entities;
using System.Collections.Generic;
using System.Windows.Input;

namespace KPLN_Parameters_Ribbon.Forms.ViewModels
{
    public sealed class SumParametersVM
    {
        public SumParametersVM(UIApplication uiapp)
        {
            CurrentSumParametersM = new SumParametersM(uiapp.ActiveUIDocument.Document);
            SetUserSelection(uiapp.ActiveUIDocument.Document, SumParametersM.GetUserSelection(uiapp));

            UpdateSumResultsCmd = new RelayCommand<object>(_ => UpdateSumResults());
            IncreaseRoundDigitsCmd = new RelayCommand<object>(_ => CurrentSumParametersM.IncrementRoundDigits());
            DecreaseRoundDigitsCmd = new RelayCommand<object>(_ => CurrentSumParametersM.DecrementRoundDigits());
            CloseWindowCmd = new RelayCommand<object>(CloseWindow);
        }

        public SumParametersM CurrentSumParametersM { get; set; }

        public ICommand UpdateSumResultsCmd { get; }

        public ICommand IncreaseRoundDigitsCmd { get; }

        public ICommand DecreaseRoundDigitsCmd { get; }

        public ICommand CloseWindowCmd { get; }

        public void UpdateSumResults() =>
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandUpdateParameterSums(CurrentSumParametersM));

        public void SetUserSelection(Document doc, IEnumerable<Element> userSelElems) =>
            CurrentSumParametersM.SetUserSelection(doc, userSelElems);

        public void CloseWindow(object windObj)
        {
            if (windObj is System.Windows.Window window)
                window.Close();
        }
    }
}
