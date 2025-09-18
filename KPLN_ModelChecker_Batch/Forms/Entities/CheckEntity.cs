using Autodesk.Revit.DB;
using KPLN_ModelChecker_Lib;
using KPLN_ModelChecker_Lib.Core;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_ModelChecker_Batch.Forms.Entities
{
    public sealed class CheckEntity : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private bool _isChecked = false;

        public CheckEntity(AbstrCheck abstrCheck)
        {
            CurrentAbstrCheck = abstrCheck;
            Name = abstrCheck.PluginName;
        }

        public string Name { get; }

        public bool IsChecked
        {
            get => _isChecked;
            set
            {
                _isChecked = value;
                OnPropertyChanged();
            }
        }

        public AbstrCheck CurrentAbstrCheck { get; }

        public CheckerEntity[] RunCommand()
        {
            Element[] elemsToCheck = CurrentAbstrCheck.GetElemsToCheck();
            return CurrentAbstrCheck.ExecuteCheck(elemsToCheck, false);
        }

        private void OnPropertyChanged([CallerMemberName] string name = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
