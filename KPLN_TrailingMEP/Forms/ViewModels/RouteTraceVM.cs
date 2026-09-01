using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using KPLN_TrailingMEP.ExecutableCommand;
using KPLN_TrailingMEP.Forms.Commands;
using KPLN_TrailingMEP.Forms.Entities;
using System;
using System.Windows;
using System.Windows.Input;

namespace KPLN_TrailingMEP.Forms.ViewModels
{
    public sealed class RouteTraceVM : IDisposable
    {
        private readonly RouteTraceForm _mainWindow;
        private readonly UIApplication _uiapp;
        private bool _isDisposed;

        public RouteTraceVM(RouteTraceForm mainWindow, UIApplication uiapp)
        {
            _mainWindow = mainWindow;
            _uiapp = uiapp;
            CurrentRouteTraceM = new RouteTraceM(uiapp);
            CurrentRouteTraceM.RouteSettingsChanged += (_, __) => RebuildPreviewRouteBySettings();
            _uiapp.Application.DocumentChanged += OnDocumentChanged;

            PickBundleCmd = new RelayCommand<object>(_ => PickBundle());
            PickTargetPointCmd = new RelayCommand<object>(_ => PickTargetPoint());
            CreatePreviewRouteCmd = new RelayCommand<object>(_ => CreatePreviewRoute());
            PickPreviewRouteCmd = new RelayCommand<object>(_ => PickPreviewRoute());
            DeletePreviewRouteCmd = new RelayCommand<object>(_ => DeletePreviewRoute());
            BuildRouteCmd = new RelayCommand<object>(_ => BuildRoute());
            CloseWindowCmd = new RelayCommand<object>(CloseWindow);
        }

        public RouteTraceM CurrentRouteTraceM { get; set; }

        public ICommand PickBundleCmd { get; }

        public ICommand PickTargetPointCmd { get; }

        public ICommand CreatePreviewRouteCmd { get; }

        public ICommand PickPreviewRouteCmd { get; }

        public ICommand DeletePreviewRouteCmd { get; }

        public ICommand BuildRouteCmd { get; }

        public ICommand CloseWindowCmd { get; }

        public void PickBundle() => KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new PickBundleExcCmd(CurrentRouteTraceM, _mainWindow));

        public void PickTargetPoint() => KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new PickTargetPointExcCmd(CurrentRouteTraceM, _mainWindow));

        public void CreatePreviewRoute() => KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CreatePreviewRouteExcCmd(CurrentRouteTraceM));

        public void PickPreviewRoute() => KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new PickPreviewRouteExcCmd(CurrentRouteTraceM, _mainWindow));

        public void DeletePreviewRoute() => KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new DeletePreviewRouteExcCmd(CurrentRouteTraceM));

        public void BuildRoute() => KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new BuildRouteExcCmd(CurrentRouteTraceM));

        private void RebuildPreviewRouteBySettings()
        {
            if (CurrentRouteTraceM.CanCreatePreview && !CurrentRouteTraceM.IsPreviewRouteChanged)
                CreatePreviewRoute();
        }

        public void CloseWindow(object windObj)
        {
            if (windObj is Window window)
                window.Close();
        }

        public void Dispose()
        {
            if (_isDisposed)
                return;

            _uiapp.Application.DocumentChanged -= OnDocumentChanged;
            _isDisposed = true;
        }

        private void OnDocumentChanged(object sender, DocumentChangedEventArgs args)
        {
            if (_isDisposed || CurrentRouteTraceM.IsRouteChangeTrackingSuspended)
                return;

            Document changedDocument = args.GetDocument();
            if (changedDocument == null || !changedDocument.Equals(CurrentRouteTraceM.Doc))
                return;

            if (!CurrentRouteTraceM.HasPreviewRouteChanges(args.GetModifiedElementIds(), args.GetDeletedElementIds()))
                return;

            _mainWindow.Dispatcher.BeginInvoke(new Action(CurrentRouteTraceM.MarkPreviewRouteChanged));
        }
    }
}
