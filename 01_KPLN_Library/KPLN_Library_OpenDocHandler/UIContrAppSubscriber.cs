using Autodesk.Revit.UI;
using KPLN_Library_OpenDocHandler.Core;
using NLog;
using System;

namespace KPLN_Library_OpenDocHandler
{
    public sealed class UIContrAppSubscriber : IDisposable
    {
        public UIContrAppSubscriber(UIControlledApplication application, Logger logger, IFieldChangedNotifier service)
        {
            CurrentUIContrApp = application;
            CurrentOpenDocHandler = new OpenDocHandler(logger, service);

            CurrentUIContrApp.DialogBoxShowing += CurrentOpenDocHandler.OnDialogBoxShowing;
            CurrentUIContrApp.ControlledApplication.DocumentOpening += CurrentOpenDocHandler.OnDocumentOpening;
            CurrentUIContrApp.ControlledApplication.DocumentClosing += CurrentOpenDocHandler.OnDocumentClosing;
            CurrentUIContrApp.ControlledApplication.DocumentClosed += CurrentOpenDocHandler.OnDocumentClosed;
            CurrentUIContrApp.ControlledApplication.FailuresProcessing += CurrentOpenDocHandler.OnFailureProcessing;
        }

        internal OpenDocHandler CurrentOpenDocHandler { get; private set; }

        internal UIControlledApplication CurrentUIContrApp { get; private set; }

        public void Dispose()
        {
            CurrentUIContrApp.DialogBoxShowing -= CurrentOpenDocHandler.OnDialogBoxShowing;
            CurrentUIContrApp.ControlledApplication.DocumentOpening -= CurrentOpenDocHandler.OnDocumentOpening;
            CurrentUIContrApp.ControlledApplication.DocumentClosed -= CurrentOpenDocHandler.OnDocumentClosed;
            CurrentUIContrApp.ControlledApplication.FailuresProcessing -= CurrentOpenDocHandler.OnFailureProcessing;


            GC.SuppressFinalize(this);
        }
    }
}
