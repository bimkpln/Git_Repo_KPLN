using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Parameters_Ribbon.ExternalEventHandler;
using System;
using System.Windows;

namespace KPLN_Parameters_Ribbon.Forms.Common
{
    internal static class FormEventSubscriptionHelper
    {
#if !Debug2020 && !Revit2020
        public static ExternalEvent CreateSelectionChangedEvent(Action<SelectionChangedHandler> configureHandler)
        {
            SelectionChangedHandler selectionHandler = new SelectionChangedHandler();
            configureHandler?.Invoke(selectionHandler);

            return ExternalEvent.Create(selectionHandler);
        }

        public static ExternalEvent CreateSelectionUnsubscribeEvent(EventHandler<SelectionChangedEventArgs> handler)
        {
            UnsubSelectionChangedHandler unsubSelHandler = new UnsubSelectionChangedHandler() { Handler = handler };
            return ExternalEvent.Create(unsubSelHandler);
        }

        public static void SubscribeSelectionChanged(UIApplication uiapp, Window window, EventHandler<SelectionChangedEventArgs> handler, ExternalEvent unsubEvent)
        {
            uiapp.SelectionChanged += handler;
            window.Closed += (s, e) => unsubEvent.Raise();
        }
#endif
    }
}
