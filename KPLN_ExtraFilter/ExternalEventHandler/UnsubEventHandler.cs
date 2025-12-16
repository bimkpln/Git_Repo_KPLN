using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using System;

namespace KPLN_ExtraFilter.ExternalEventHandler
{
    public sealed class UnsubEventHandler : IExternalEventHandler
    {
        public string GetName() => "UnsubEventHandler";

#if Debug2020 || Revit2020
        public void Execute(UIApplication app) => throw new NotImplementedException();
#else
        public EventHandler<SelectionChangedEventArgs> Handler { get; set; }

        public void Execute(UIApplication app) => app.SelectionChanged -= Handler;
#endif
    }
}
