using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using System;

namespace KPLN_ExtraFilter.ExternalEventHandler
{
    public sealed class UnsubViewActHandler : IExternalEventHandler
    {
        public string GetName() => "UnsubViewActHandler";

        public EventHandler<ViewActivatedEventArgs> Handler { get; set; }

        public void Execute(UIApplication app) => app.ViewActivated -= Handler;
    }
}
