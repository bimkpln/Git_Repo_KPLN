using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using System;

namespace KPLN_ExtraFilter.ExternalEventHandler
{
    public sealed class SelectionByModelSubscriptionHandler : IExternalEventHandler
    {
        public EventHandler<ViewActivatedEventArgs> ViewHandler { get; set; }

        public bool ShouldSubscribeViewChanged { get; set; }

        private bool _isViewChangedSubscribed;

#if !Debug2020 && !Revit2020
        public EventHandler<SelectionChangedEventArgs> SelectionHandler { get; set; }

        public bool ShouldSubscribeSelectionChanged { get; set; }

        private bool _isSelectionChangedSubscribed;
#endif

        public void Execute(UIApplication app)
        {
            SetViewChangedSubscription(app);

#if !Debug2020 && !Revit2020
            SetSelectionChangedSubscription(app);
#endif
        }

        public string GetName() => "SelectionByModelSubscriptionHandler";

        private void SetViewChangedSubscription(UIApplication app)
        {
            if (app == null || ViewHandler == null)
                return;

            if (ShouldSubscribeViewChanged && !_isViewChangedSubscribed)
            {
                app.ViewActivated += ViewHandler;
                _isViewChangedSubscribed = true;
                return;
            }

            if (!ShouldSubscribeViewChanged && _isViewChangedSubscribed)
            {
                app.ViewActivated -= ViewHandler;
                _isViewChangedSubscribed = false;
            }
        }

#if !Debug2020 && !Revit2020
        private void SetSelectionChangedSubscription(UIApplication app)
        {
            if (app == null || SelectionHandler == null)
                return;

            if (ShouldSubscribeSelectionChanged && !_isSelectionChangedSubscribed)
            {
                app.SelectionChanged += SelectionHandler;
                _isSelectionChangedSubscribed = true;
                return;
            }

            if (!ShouldSubscribeSelectionChanged && _isSelectionChangedSubscribed)
            {
                app.SelectionChanged -= SelectionHandler;
                _isSelectionChangedSubscribed = false;
            }
        }
#endif
    }
}