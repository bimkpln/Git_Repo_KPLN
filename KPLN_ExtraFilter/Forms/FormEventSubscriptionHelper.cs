using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_ExtraFilter.ExternalEventHandler;
using System;
using System.Windows;

namespace KPLN_ExtraFilter.Forms
{
    internal static class FormEventSubscriptionHelper
    {
        public static ExternalEvent CreateViewChangedEvent(Action<ViewActivatedHandler> configureHandler)
        {
            ViewActivatedHandler viewHandler = new ViewActivatedHandler();
            configureHandler?.Invoke(viewHandler);

            return ExternalEvent.Create(viewHandler);
        }

        public static ExternalEvent CreateViewUnsubscribeEvent(EventHandler<ViewActivatedEventArgs> handler)
        {
            UnsubViewActHandler unsubViewHandler = new UnsubViewActHandler() { Handler = handler };
            return ExternalEvent.Create(unsubViewHandler);
        }

        public static void SubscribeViewChanged(UIApplication uiapp, Window window, EventHandler<ViewActivatedEventArgs> handler, ExternalEvent unsubEvent)
        {
            uiapp.ViewActivated += handler;
            window.Closed += (s, e) => unsubEvent.Raise();
        }

#if !Debug2020 && !Revit2020
        public static ExternalEvent CreateSelectionChangedEvent(Action<SelectionChangedHandler> configureHandler)
        {
            SelectionChangedHandler selectionHandler = new SelectionChangedHandler();
            configureHandler?.Invoke(selectionHandler);

            return ExternalEvent.Create(selectionHandler);
        }

        public static ExternalEvent CreateSelectionUnsubscribeEvent(EventHandler<SelectionChangedEventArgs> handler)
        {
            UnsubEventHandler unsubSelHandler = new UnsubEventHandler() { Handler = handler };
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
