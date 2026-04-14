using Autodesk.Revit.UI;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_Clashes_Ribbon.ExternalEventHandler
{
    public sealed class ViewActivatedHandler : IExternalEventHandler, INotifyPropertyChanged
    {
        private UIApplication _currnetUIApplication;

        public UIApplication CurrnetUIApplication 
        {
            get => _currnetUIApplication;
            set
            {
                _currnetUIApplication = value;
                NotifyPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public void Execute(UIApplication app)
        {
            CurrnetUIApplication = app;
        }

        public string GetName() => "ViewActivatedHandler";

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "") =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
