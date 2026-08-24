using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using System;

namespace KPLN_Parameters_Ribbon.ExternalEventHandler
{
    public sealed class UnsubSelectionChangedHandler : IExternalEventHandler
    {
        public string GetName() => "UnsubSelectionChangedHandler";

#if Debug2020 || Revit2020
        public void Execute(UIApplication app) => throw new NotImplementedException();
#else
        public EventHandler<SelectionChangedEventArgs> Handler { get; set; }

        public void Execute(UIApplication app) => app.SelectionChanged -= Handler;
#endif
    }
}
